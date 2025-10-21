#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Teams Link Scheduler (COM) - v1.2
# - FIX: Task Scheduler folder path. Use "\TeamsLinks" (single leading slash) and
#   create folders step-by-step to avoid 0x8007007B (invalid name).

import os, sys, re, traceback
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime, timezone
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH = os.path.join(BASE_DIR, "logo.png")   # 로고 파일명
APP_TITLE = "Teams Link Scheduler (COM)"
TASK_FOLDER = r"\TeamsLinks"   # <-- single leading slash only
DEFAULT_TASK_PREFIX = ""
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")
SCHEDULE_LABELS = {"ONCE": "1回", "DAILY": "毎日", "WEEKLY": "毎週"}
SCHEDULE_FROM_LABEL = {v: k for k, v in SCHEDULE_LABELS.items()}

TASK_TRIGGER_TIME, TASK_TRIGGER_DAILY, TASK_TRIGGER_WEEKLY = 1,2,3
TASK_ACTION_EXEC = 0
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_CREATE_OR_UPDATE = 6
TASK_RUNLEVEL_LUA, TASK_RUNLEVEL_HIGHEST = 0,1

DAYS = {"SUN":1, "MON":2, "TUE":4, "WED":8, "THU":16, "FRI":32, "SAT":64}
WEEKDAY_MAP = [("月","MON"),("火","TUE"),("水","WED"),("木","THU"),("金","FRI"),("土","SAT"),("日","SUN")]

def extract_url_from_cmdargs(args: str) -> str:
    # act.Arguments 예: /c start "" "ms-teams://..."  또는  /c start "" "https://..."
    m = re.search(r'"(ms-teams://[^"]+|https?://[^"]+)"\s*$', args or "")
    return m.group(1) if m else ""

def apply_task_to_form(self, task_name: str):
    info = get_task_info(task_name)

    # 기본 값들
    self.name_var.set(info.get("name", ""))
    self.url_var.set(info.get("url", ""))

    # 빈번도(일본어 라벨 반영)
    sched_key = info.get("schedule", "ONCE")
    if 'SCHEDULE_LABELS' in globals():
        self.schedule_var.set(SCHEDULE_LABELS.get(sched_key, sched_key))
    else:
        self.schedule_var.set(sched_key)

    # 날짜/시간
    self.date_var.set(info.get("start_date", ""))
    self.time_var.set(info.get("start_time", ""))

    # 요일 체크 (weekly일 때만)
    weekdays = set(info.get("weekdays", []))
    for code, var in self.weekly_vars.items():
        var.set(code in weekdays)

    # WEEKLY가 아니면 요일 전부 false로 정리(표시 일관성)
    if sched_key != "WEEKLY":
        for var in self.weekly_vars.values():
            var.set(False)

    # 전 항목 비활성화
    self._lock_all_inputs()

    self.status.set(f"選択中: {task_name}")

def debug(*a):
    if os.environ.get("DEBUG") == "1":
        print("[DEBUG]", *a, file=sys.stderr)

def require_pywin32():
    if win32com is None:
        messagebox.showerror("エラー", "pywin32 が見つかりません。install_env.bat を実行してください。")
        raise SystemExit(1)

def connect_service():
    svc = win32com.client.Dispatch("Schedule.Service"); svc.Connect(); return svc

def get_or_create_folder(svc, path):
    # robust folder creation: walk \A\B\C
    root = svc.GetFolder("\\")
    current = "\\"
    # normalize
    parts = [p for p in path.split("\\") if p]
    for p in parts:
        next_path = (current.rstrip("\\") + "\\" + p)
        try:
            svc.GetFolder(next_path)
            debug("Folder exists:", next_path)
        except Exception:
            debug("Creating folder:", next_path)
            root.CreateFolder(next_path, "")
        current = next_path
    return svc.GetFolder(current)


def make_taskdef(svc, runlevel=0):
    td = svc.NewTask(0)
    td.RegistrationInfo.Description = "Created by Python (win32com)"
    td.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
    td.Principal.RunLevel = runlevel
    s = td.Settings
    s.Enabled = True; s.StartWhenAvailable=True; s.AllowDemandStart=True
    s.DisallowStartIfOnBatteries=False; s.StopIfGoingOnBatteries=False
    s.Hidden=False; s.ExecutionTimeLimit="PT0S"
    return td

def register_task(schedule, name, url, date_str, time_str, weekdays_codes, run_as_admin=False):
    require_pywin32()
    if not url.startswith(("http://","https://","ms-teams://")): raise ValueError("URL は http(s):// または ms-teams:// で。")
    if not TIME_RE.match(time_str or ""): raise ValueError("時刻は HH:MM。")
    if date_str and not DATE_RE.match(date_str): raise ValueError("日付は YYYY-MM-DD。")
    if schedule=="WEEKLY" and not weekdays_codes: raise ValueError("WEEKLY は曜日を選択。")

    start_boundary = f"{date_str}T{time_str}:00" if date_str else datetime.now().strftime(f"%Y-%m-%dT{time_str}:00")
    cmd_path = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), r"System32\cmd.exe")
    debug("Params", dict(schedule=schedule,name=name,url=url,start=start_boundary,days=weekdays_codes,admin=run_as_admin,cmd=cmd_path))

    try:
        svc = connect_service(); folder = get_or_create_folder(svc, TASK_FOLDER)
        td = make_taskdef(svc, TASK_RUNLEVEL_HIGHEST if run_as_admin else TASK_RUNLEVEL_LUA)

        if schedule=="ONCE":
            trig = td.Triggers.Create(TASK_TRIGGER_TIME); trig.StartBoundary = start_boundary
        elif schedule=="DAILY":
            trig = td.Triggers.Create(TASK_TRIGGER_DAILY); trig.StartBoundary=start_boundary; trig.DaysInterval=1
        elif schedule=="WEEKLY":
            trig = td.Triggers.Create(TASK_TRIGGER_WEEKLY); trig.StartBoundary=start_boundary; trig.WeeksInterval=1
            mask=0
            for w in weekdays_codes: mask |= DAYS[w]
            trig.DaysOfWeek = mask
        else: raise ValueError("Invalid schedule")

        act = td.Actions.Create(TASK_ACTION_EXEC); act.Path = cmd_path; act.Arguments = f'/c start "" "{url}"'
        folder.RegisterTaskDefinition(name, td, TASK_CREATE_OR_UPDATE, "", "", TASK_LOGON_INTERACTIVE_TOKEN, "")
        debug("Registered", f"{TASK_FOLDER}\{name}")
    except Exception as e:
        if pythoncom and isinstance(e, pythoncom.com_error):
            print("COM ERROR:", hex(e.hresult), e.excepinfo, file=sys.stderr)
        traceback.print_exc(); raise

def list_tasks():
    require_pywin32(); svc = connect_service()
    try: folder = svc.GetFolder(TASK_FOLDER)
    except Exception: return []
    return [(t.Name, str(t.NextRunTime), int(t.State)) for t in folder.GetTasks(0)]

def delete_task(task_name):
    require_pywin32(); svc=connect_service(); folder=get_or_create_folder(svc, TASK_FOLDER); folder.DeleteTask(task_name,0)

def run_task_now(task_name):
    require_pywin32(); svc=connect_service(); folder=get_or_create_folder(svc, TASK_FOLDER); folder.GetTask(task_name).Run("")

def _parse_start_boundary(sb: str):
    """'2025-10-21T09:55:00' / '...Z' / '...+09:00'"""
    if not sb:
        return "", ""
    s = sb.replace("Z", "+00:00") if sb.endswith("Z") else sb
    try:
        dt = datetime.fromisoformat(s)
    except ValueError:
        try:
            dt = datetime.fromisoformat(s.split(".")[0])
        except Exception:
            return "", ""
    return dt.strftime("%Y-%m-%d"), dt.strftime("%H:%M")

def _days_mask_to_codes(mask: int):
    """Task Scheduler の DaysOfWeek → ['MON','WED',...]"""
    out = []
    for code, bit in DAYS.items():
        if mask & bit:
            out.append(code)
    return out

def get_task_info(task_name: str):
    """タスク名, URL, スケジュール種別, 開始日, 時刻, 曜日(weekly時) を返す"""
    require_pywin32()
    svc = connect_service()
    folder = get_or_create_folder(svc, TASK_FOLDER)
    t = folder.GetTask(task_name)
    td = t.Definition

    # URL
    try:
        act = td.Actions.Item(1)
    except Exception:
        act = next(iter(td.Actions), None)
    url = extract_url_from_cmdargs(getattr(act, "Arguments", "")) if act else ""

    # Trigger
    try:
        trig = td.Triggers.Item(1)
    except Exception:
        trig = next(iter(td.Triggers), None)

    trig_type = int(getattr(trig, "Type", TASK_TRIGGER_TIME) or TASK_TRIGGER_TIME)
    schedule_key = {
        TASK_TRIGGER_TIME: "ONCE",
        TASK_TRIGGER_DAILY: "DAILY",
        TASK_TRIGGER_WEEKLY: "WEEKLY"
    }.get(trig_type, "ONCE")

    start_date, start_time = _parse_start_boundary(getattr(trig, "StartBoundary", "") or "")

    weekdays_codes = []
    if schedule_key == "WEEKLY":
        try:
            mask = int(getattr(trig, "DaysOfWeek", 0) or 0)
        except Exception:
            mask = 0
        weekdays_codes = _days_mask_to_codes(mask)

    return {
        "name": t.Name,
        "url": url,
        "schedule": schedule_key,
        "start_date": start_date,
        "start_time": start_time,
        "weekdays": weekdays_codes
    }


class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_TITLE); self.geometry("840x660"); self.minsize(780,560)
        self.logo_img = None
        if os.path.exists(LOGO_PATH):
            try:
                self.logo_img = tk.PhotoImage(file=LOGO_PATH)
                try:
                    self.iconphoto(True, self.logo_img)
                except Exception:
                    pass
                top = ttk.Frame(self)
                top.pack(fill="x", padx=8, pady=8)
                ttk.Label(top, image=self.logo_img).pack()
            except Exception as e:
                print("[Logo] load failed:", e, file=sys.stderr)
        self.create_widgets(); self.refresh_tasks()
    def create_widgets(self):
        pad={"padx":8,"pady":6}; frm=ttk.LabelFrame(self,text="スケジュール作成"); frm.pack(fill="x",**pad)
        ttk.Label(frm,text="タスク名").grid(row=0,column=0,sticky="w",**pad); self.name_var=tk.StringVar(value=DEFAULT_TASK_PREFIX)
        self.entry_name = ttk.Entry(frm,textvariable=self.name_var,width=36)
        self.entry_name.grid(row=0,column=1,sticky="w",**pad)
        ttk.Label(frm,text="Teamsリンク（URL）").grid(row=1,column=0,sticky="w",**pad); self.url_var=tk.StringVar()
        self.entry_url = ttk.Entry(frm,textvariable=self.url_var,width=70)
        self.entry_url.grid(row=1,column=1,columnspan=3,sticky="we",**pad)
        ttk.Label(frm,text="頻度").grid(row=2,column=0,sticky="w",**pad); self.schedule_var = tk.StringVar(value=SCHEDULE_LABELS["WEEKLY"])
        self.combo_schedule = ttk.Combobox(frm,textvariable=self.schedule_var,state="readonly",width=10,values=list(SCHEDULE_LABELS.values()))
        self.combo_schedule.grid(row=2,column=1,sticky="w",**pad)
        
        ttk.Label(frm,text="開始日（YYYY-MM-DD）").grid(row=2,column=2,sticky="e",**pad); self.date_var=tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        self.entry_date = ttk.Entry(frm,textvariable=self.date_var,width=14)
        self.entry_date.grid(row=2,column=3,sticky="w",**pad)
        ttk.Label(frm,text="時刻（HH:MM 24h）").grid(row=3,column=2,sticky="e",**pad); self.time_var=tk.StringVar(value="09:55")
        self.entry_time = ttk.Entry(frm,textvariable=self.time_var,width=10)
        self.entry_time.grid(row=3,column=3,sticky="w",**pad)
        self.weekly_vars={}; 
        self.days_frame=ttk.Frame(frm); 
        self.days_frame.grid(row=3,column=1,sticky="w",**pad)
        for i, (label, code) in enumerate([("月","MON"),("火","TUE"),("水","WED"),("木","THU"),("金","FRI"),("土","SAT"),("日","SUN")]):
            v = tk.BooleanVar(value=(code in ["MON","TUE","WED","THU","FRI"]))
            self.weekly_vars[code] = v
            ttk.Checkbutton(self.days_frame, text=label, variable=v).grid(row=0, column=i, sticky="w")
        # self.admin_var=tk.BooleanVar(value=False); ttk.Checkbutton(frm,text="管理者権限で実行（要：Pythonを管理者で起動）",variable=self.admin_var).grid(row=4,column=1,columnspan=3,sticky="w",**pad)
        btn_fr=ttk.Frame(frm); btn_fr.grid(row=5,column=0,columnspan=4,sticky="e",**pad)
        ttk.Button(btn_fr,text="作成/更新",command=self.on_create).pack(side="left",padx=4); ttk.Button(btn_fr,text="再読み込み",command=self.refresh_tasks).pack(side="left",padx=4)
        list_fr=ttk.LabelFrame(self,text="登録済みタスク（\TeamsLinks）"); list_fr.pack(fill="both",expand=True,**pad)
        cols=("TaskName","NextRun","State"); self.tree=ttk.Treeview(list_fr,columns=cols,show="headings",height=12)
        for c,t,w in [("TaskName","タスク名",420),("NextRun","次回実行",180),("State","状態",80)]: self.tree.heading(c,text=t); self.tree.column(c,width=w)
        self.tree.pack(fill="both",expand=True,padx=6,pady=6); ctl_fr=ttk.Frame(list_fr); ctl_fr.pack(fill="x",padx=6,pady=6)
        self.tree.bind("<Return>", self.on_tree_row_activate)
        self.AUTO_EDIT_ON_SINGLE_CLICK = True
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_row_select)
        ttk.Button(ctl_fr,text="選択削除",command=self.on_delete).pack(side="left"); 
        ttk.Button(ctl_fr,text="実行",command=self.on_run).pack(side="left",padx=6)
        self.status=tk.StringVar(value="準備完了"); ttk.Label(self,textvariable=self.status,anchor="w").pack(fill="x")
    def on_create(self):
        try:
            name = self.name_var.get().strip() or DEFAULT_TASK_PREFIX + "Task"
            for ch in '\\/:*?"<>|':
                name = name.replace(ch, "_")

            selected_label = self.schedule_var.get()
            schedule_key = SCHEDULE_FROM_LABEL.get(selected_label, selected_label) 

            weekdays = (
                [code for code, v in self.weekly_vars.items() if v.get()]
                if schedule_key == "WEEKLY" else []
            )

            register_task(
                schedule_key,
                name,
                self.url_var.get().strip(),
                self.date_var.get().strip(),
                self.time_var.get().strip(),
                weekdays
            )
            messagebox.showinfo("完了", f"タスクを作成/更新しました: {TASK_FOLDER}\\{name}")
            self.refresh_tasks()
        except Exception as e: messagebox.showerror("エラー", str(e))

    def refresh_tasks(self):
        for i in self.tree.get_children(): self.tree.delete(i)
        try: rows=list_tasks()
        except Exception as e: self.status.set(f"一覧取得に失敗: {e}"); return
        for name,nextrun,state in rows: self.tree.insert("", "end", values=(name,nextrun,state))
        self.status.set(f"{len(rows)}件のタスク")

    def on_delete(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("注意","削除するタスクを選択してください。"); return
        name=self.tree.item(sel[0],"values")[0]
        try: delete_task(name); self.refresh_tasks(); messagebox.showinfo("完了","削除しました。")
        except Exception as e: messagebox.showerror("エラー", str(e))

    def on_run(self):
        sel=self.tree.selection()
        if not sel: messagebox.showwarning("注意","実行するタスクを選択してください。"); return
        name=self.tree.item(sel[0],"values")[0]
        try: run_task_now(name); messagebox.showinfo("実行","実行要求を送信しました。")
        except Exception as e: messagebox.showerror("エラー", str(e))

    def on_tree_row_activate(self, event=None):
        try:
            item_id = self.tree.identify_row(event.y) if event and hasattr(event, "y") else None
            if not item_id:
                # selection fallback
                sel = self.tree.selection()
                if not sel:
                    messagebox.showwarning("注意", "反映するタスクを選択してください。")
                    return
                item_id = sel[0]

            task_name = self.tree.item(item_id, "values")[0]
            self.apply_task_to_form(task_name)
        except Exception as e:
            messagebox.showerror("エラー", f"タスク情報の反映に失敗しました: {e}")

    def on_tree_row_select(self, event=None):
        if not getattr(self, "AUTO_EDIT_ON_SINGLE_CLICK", False):
            return
        try:
            sel = self.tree.selection()
            if not sel:
                return
            task_name = self.tree.item(sel[0], "values")[0]
            apply_task_to_form(self, task_name)
        except Exception as e:
            self.status.set(f"反映失敗: {e}")
    def _set_weekday_checks_enabled(self, enabled: bool):
        state = ("!disabled" if enabled else "disabled")
        for w in self.days_frame.winfo_children():
            try:
                w.state((state,))
            except Exception:
                try:
                    w.configure(state="normal" if enabled else "disabled")
                except Exception:
                    pass

    def _lock_all_inputs(self):
        self.entry_name.configure(state="disabled")
        self.entry_url.configure(state="disabled")
        self.entry_date.configure(state="disabled")
        self.entry_time.configure(state="disabled")
        self.combo_schedule.configure(state="disabled")
        self._set_weekday_checks_enabled(False)

    def _unlock_all_inputs(self):
        self.entry_name.configure(state="normal")
        self.entry_url.configure(state="normal")
        self.entry_date.configure(state="normal")
        self.entry_time.configure(state="normal")
        self.combo_schedule.configure(state="readonly")
        label = self.schedule_var.get()
        key = SCHEDULE_FROM_LABEL.get(label, label)
        self._set_weekday_checks_enabled(key == "WEEKLY")
if __name__ == "__main__":
    if os.name != "nt": print("Windows only."); sys.exit(1)
    app=App(); app.mainloop()
