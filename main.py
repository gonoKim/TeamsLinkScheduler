#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Teams Link Scheduler (COM) - v1.2
# - FIX: Task Scheduler folder path. Use "\TeamsLinks" (single leading slash) and
#   create folders step-by-step to avoid 0x8007007B (invalid name).

import os, sys, re, traceback
import tkinter as tk
from tkinter import ttk, messagebox
from datetime import datetime
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

APP_TITLE = "Teams Link Scheduler (COM)"
TASK_FOLDER = r"\TeamsLinks"   # <-- single leading slash only
DEFAULT_TASK_PREFIX = ""
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

TASK_TRIGGER_TIME, TASK_TRIGGER_DAILY, TASK_TRIGGER_WEEKLY = 1,2,3
TASK_ACTION_EXEC = 0
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_CREATE_OR_UPDATE = 6
TASK_RUNLEVEL_LUA, TASK_RUNLEVEL_HIGHEST = 0,1

DAYS = {"SUN":1, "MON":2, "TUE":4, "WED":8, "THU":16, "FRI":32, "SAT":64}
WEEKDAY_MAP = [("月","MON"),("火","TUE"),("水","WED"),("木","THU"),("金","FRI"),("土","SAT"),("日","SUN")]

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

class App(tk.Tk):
    def __init__(self):
        super().__init__(); self.title(APP_TITLE); self.geometry("840x660"); self.minsize(780,560)
        self.create_widgets(); self.refresh_tasks()
    def create_widgets(self):
        pad={"padx":8,"pady":6}; frm=ttk.LabelFrame(self,text="スケジュール作成"); frm.pack(fill="x",**pad)
        ttk.Label(frm,text="タスク名").grid(row=0,column=0,sticky="w",**pad); self.name_var=tk.StringVar(value=DEFAULT_TASK_PREFIX)
        ttk.Entry(frm,textvariable=self.name_var,width=36).grid(row=0,column=1,sticky="w",**pad)
        ttk.Label(frm,text="Teamsリンク（URL）").grid(row=1,column=0,sticky="w",**pad); self.url_var=tk.StringVar()
        ttk.Entry(frm,textvariable=self.url_var,width=70).grid(row=1,column=1,columnspan=3,sticky="we",**pad)
        ttk.Label(frm,text="頻度").grid(row=2,column=0,sticky="w",**pad); self.schedule_var=tk.StringVar(value="WEEKLY")
        ttk.Combobox(frm,textvariable=self.schedule_var,state="readonly",width=10,values=["ONCE","DAILY","WEEKLY"]).grid(row=2,column=1,sticky="w",**pad)
        ttk.Label(frm,text="開始日（YYYY-MM-DD）").grid(row=2,column=2,sticky="e",**pad); self.date_var=tk.StringVar(value=datetime.now().strftime("%Y-%m-%d"))
        ttk.Entry(frm,textvariable=self.date_var,width=14).grid(row=2,column=3,sticky="w",**pad)
        ttk.Label(frm,text="時刻（HH:MM 24h）").grid(row=3,column=2,sticky="e",**pad); self.time_var=tk.StringVar(value="09:55")
        ttk.Entry(frm,textvariable=self.time_var,width=10).grid(row=3,column=3,sticky="w",**pad)
        self.weekly_vars={}; days_frame=ttk.Frame(frm); days_frame.grid(row=3,column=1,sticky="w",**pad)
        for i,(label,code) in enumerate([("月","MON"),("火","TUE"),("水","WED"),("木","THU"),("金","FRI"),("土","SAT"),("日","SUN")]):
            v=tk.BooleanVar(value=(code in ["MON","TUE","WED","THU","FRI"])); self.weekly_vars[code]=v; ttk.Checkbutton(days_frame,text=label,variable=v).grid(row=0,column=i,sticky="w")
        self.admin_var=tk.BooleanVar(value=False); ttk.Checkbutton(frm,text="管理者権限で実行（要：Pythonを管理者で起動）",variable=self.admin_var).grid(row=4,column=1,columnspan=3,sticky="w",**pad)
        btn_fr=ttk.Frame(frm); btn_fr.grid(row=5,column=0,columnspan=4,sticky="e",**pad)
        ttk.Button(btn_fr,text="作成/更新",command=self.on_create).pack(side="left",padx=4); ttk.Button(btn_fr,text="再読み込み",command=self.refresh_tasks).pack(side="left",padx=4)
        list_fr=ttk.LabelFrame(self,text="登録済みタスク（\TeamsLinks）"); list_fr.pack(fill="both",expand=True,**pad)
        cols=("TaskName","NextRun","State"); self.tree=ttk.Treeview(list_fr,columns=cols,show="headings",height=12)
        for c,t,w in [("TaskName","タスク名",420),("NextRun","次回実行",180),("State","状態",80)]: self.tree.heading(c,text=t); self.tree.column(c,width=w)
        self.tree.pack(fill="both",expand=True,padx=6,pady=6); ctl_fr=ttk.Frame(list_fr); ctl_fr.pack(fill="x",padx=6,pady=6)
        ttk.Button(ctl_fr,text="選択削除",command=self.on_delete).pack(side="left"); ttk.Button(ctl_fr,text="今すぐ実行",command=self.on_run).pack(side="left",padx=6)
        self.status=tk.StringVar(value="準備完了"); ttk.Label(self,textvariable=self.status,anchor="w").pack(fill="x")
    def on_create(self):
        try:
            name=self.name_var.get().strip() or DEFAULT_TASK_PREFIX+"Task"
            for ch in '\\/:*?"<>|': name=name.replace(ch,"_")
            register_task(self.schedule_var.get(), name, self.url_var.get().strip(), self.date_var.get().strip(), self.time_var.get().strip(),
                          [code for code,v in self.weekly_vars.items() if v.get()] if self.schedule_var.get()=="WEEKLY" else [],
                          self.admin_var.get())
            messagebox.showinfo("完了", f"タスクを作成/更新しました: {TASK_FOLDER}\{name}"); self.refresh_tasks()
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

if __name__ == "__main__":
    if os.name != "nt": print("Windows only."); sys.exit(1)
    app=App(); app.mainloop()
