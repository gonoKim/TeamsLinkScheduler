#!/usr/bin/env python3
# -*- coding: utf-8 -*-
# Teams Link Scheduler (COM) - v1.2
# - FIX: Task Scheduler folder path. Use "\TeamsLinks" (single leading slash) and
#   create folders step-by-step to avoid 0x8007007B (invalid name).


import os, sys, re, traceback
import ctypes, atexit
import tkinter as tk
from tkinter import ttk, messagebox, simpledialog
from datetime import datetime

# --- pywin32 読み込み --------------------------------------------------------
try:
    import win32com.client
    import pythoncom
except ImportError:
    win32com = None
    pythoncom = None

# --- 定数/設定 ---------------------------------------------------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
LOGO_PATH  = os.path.join(BASE_DIR, "logo.png")
APP_TITLE = "Link Scheduler"

# ルートフォルダ（グループの親）
TASK_FOLDER = r"\TeamsLinks"
# 既定グループ名
DEFAULT_GROUP = "default"

DEFAULT_TASK_PREFIX = ""
TIME_RE = re.compile(r"^\d{2}:\d{2}$")
DATE_RE = re.compile(r"^\d{4}-\d{2}-\d{2}$")

# UI ラベル：内部キー→表示
SCHEDULE_LABELS = {"ONCE": "1回", "DAILY": "毎日", "WEEKLY": "毎週"}
# 表示→内部キー
SCHEDULE_FROM_LABEL = {v: k for k, v in SCHEDULE_LABELS.items()}

# Task Scheduler 定数
TASK_TRIGGER_TIME, TASK_TRIGGER_DAILY, TASK_TRIGGER_WEEKLY = 1, 2, 3
TASK_ACTION_EXEC = 0
TASK_LOGON_INTERACTIVE_TOKEN = 3
TASK_CREATE_OR_UPDATE = 6
TASK_RUNLEVEL_LUA, TASK_RUNLEVEL_HIGHEST = 0, 1

# 既定入力値
DEFAULT_SCHEDULE_KEY = "WEEKLY"
DEFAULT_TIME = "09:55"
DEFAULT_WEEKDAYS = ["MON", "TUE", "WED", "THU", "FRI"]

# 曜日ビットマスク（Windows の DaysOfWeek 用）
DAYS = {"SUN":1, "MON":2, "TUE":4, "WED":8, "THU":16, "FRI":32, "SAT":64}
# 表示順（日本語→内部コード）
WEEKDAY_MAP = [("月","MON"),("火","TUE"),("水","WED"),("木","THU"),("金","FRI"),("土","SAT"),("日","SUN")]

INVALID_CHARS = r'\\/:*?"<>|'

# --- ユーティリティ ----------------------------------------------------------

def debug(*a):
    """DEBUG=1 環境変数があるとき stderr へデバッグ出力"""
    if os.environ.get("DEBUG") == "1":
        print("[DEBUG]", *a, file=sys.stderr)

def ensure_single_instance():
    """二重起動防止: 既に起動中なら既存ウィンドウを前面化して終了"""
    kernel32 = ctypes.windll.kernel32
    user32 = ctypes.windll.user32
    MUTEX_NAME = "Global\\TeamsLinkSchedulerSingleton"

    handle = kernel32.CreateMutexW(None, False, MUTEX_NAME)
    if kernel32.GetLastError() == 183:  # ERROR_ALREADY_EXISTS
        try:
            hwnd = user32.FindWindowW(None, APP_TITLE)
            if hwnd:
                SW_RESTORE = 9
                user32.ShowWindow(hwnd, SW_RESTORE)
                user32.SetForegroundWindow(hwnd)
        except Exception:
            pass
        sys.exit(0)

    atexit.register(lambda: kernel32.CloseHandle(handle))
    return handle

def today_str() -> str:
    """今日の日付 (YYYY-MM-DD) を返す"""
    return datetime.now().strftime("%Y-%m-%d")

def sanitize_name(s: str) -> str:
    """タスク名/グループ名の禁止文字を '_' に置換"""
    s = (s or "").strip()
    return re.sub(f"[{re.escape(INVALID_CHARS)}]", "_", s)

def require_pywin32():
    """pywin32 が無ければエラー表示して終了"""
    if win32com is None:
        messagebox.showerror("エラー", "pywin32 が見つかりません。install_env.bat を実行してください。")
        raise SystemExit(1)

def connect_service():
    """Task Scheduler Service に接続"""
    svc = win32com.client.Dispatch("Schedule.Service"); svc.Connect(); return svc

def get_or_create_folder(svc, path: str):
    """フォルダを階層的に作成/取得（\\A\\B\\C のように順次作成）"""
    root = svc.GetFolder("\\")
    current = "\\"
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

def group_path(group: str) -> str:
    """グループ名から完全パス（\\TeamsLinks\\group）を得る"""
    g = sanitize_name(group) or DEFAULT_GROUP
    return TASK_FOLDER.rstrip("\\") + "\\" + g

def ensure_default_group():
    """起動時に \\TeamsLinks と \\TeamsLinks\\default を必ず作成"""
    require_pywin32()
    svc = connect_service()
    get_or_create_folder(svc, TASK_FOLDER)
    get_or_create_folder(svc, group_path(DEFAULT_GROUP))

def list_groups() -> list[str]:
    """ルート直下のグループ名一覧を返す（default を含む）"""
    require_pywin32()
    svc = connect_service()
    root = get_or_create_folder(svc, TASK_FOLDER)
    groups = set()
    try:
        for f in root.GetFolders(0):
            # f.Path は \TeamsLinks\xxx
            name = (f.Path or "").rsplit("\\", 1)[-1]
            if name:
                groups.add(name)
    except Exception:
        pass
    groups.add(DEFAULT_GROUP)
    return sorted(groups, key=str.lower)

def extract_url_from_cmdargs(args: str) -> str:
    """Actions.Arguments から URL("...") を抽出"""
    m = re.search(r'"(ms-teams://[^"]+|https?://[^"]+)"\s*$', args or "")
    return m.group(1) if m else ""

def _parse_start_boundary(sb: str):
    """StartBoundary を (YYYY-MM-DD, HH:MM) に分解"""
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
    """DaysOfWeek ビットマスク→['MON','WED',...] へ"""
    out = []
    for code, bit in DAYS.items():
        if mask & bit:
            out.append(code)
    return out

# --- Task Scheduler CRUD（グループ対応） -------------------------------------

def register_task(schedule, name, url, date_str, time_str, weekdays_codes,
                  run_as_admin=False, group: str = DEFAULT_GROUP):
    """タスク登録/更新（指定グループ配下に作成）"""
    require_pywin32()

    if not url.startswith(("http://", "https://", "ms-teams://")):
        raise ValueError("URL は http(s):// または ms-teams:// で指定してください。")
    if not TIME_RE.match(time_str or ""):
        raise ValueError("時刻は HH:MM（24時間表記）で指定してください。")
    if date_str and not DATE_RE.match(date_str):
        raise ValueError("日付は YYYY-MM-DD で指定してください。")
    if schedule == "WEEKLY" and not weekdays_codes:
        raise ValueError("WEEKLY の場合は曜日を1つ以上選択してください。")

    start_boundary = f"{date_str}T{time_str}:00" if date_str \
        else datetime.now().strftime(f"%Y-%m-%dT{time_str}:00")

    cmd_path = os.path.join(os.environ.get("SystemRoot", r"C:\Windows"), r"System32\cmd.exe")
    debug("Params", dict(schedule=schedule, name=name, url=url,
                         start=start_boundary, days=weekdays_codes,
                         admin=run_as_admin, cmd=cmd_path, group=group))

    try:
        svc = connect_service()
        folder = get_or_create_folder(svc, group_path(group))  # ★グループ配下
        td = make_taskdef(svc, TASK_RUNLEVEL_HIGHEST if run_as_admin else TASK_RUNLEVEL_LUA)

        if schedule == "ONCE":
            trig = td.Triggers.Create(TASK_TRIGGER_TIME);  trig.StartBoundary = start_boundary
        elif schedule == "DAILY":
            trig = td.Triggers.Create(TASK_TRIGGER_DAILY); trig.StartBoundary = start_boundary; trig.DaysInterval = 1
        elif schedule == "WEEKLY":
            trig = td.Triggers.Create(TASK_TRIGGER_WEEKLY); trig.StartBoundary = start_boundary; trig.WeeksInterval = 1
            mask = 0
            for w in weekdays_codes:
                mask |= DAYS[w]
            trig.DaysOfWeek = mask
        else:
            raise ValueError("不正なスケジュール指定です。")

        act = td.Actions.Create(TASK_ACTION_EXEC)
        act.Path = cmd_path
        act.Arguments = f'/c start "" "{url}"'

        folder.RegisterTaskDefinition(name, td, TASK_CREATE_OR_UPDATE, "", "", TASK_LOGON_INTERACTIVE_TOKEN, "")
        debug("Registered", f"{group_path(group)}\\{name}")
    except Exception as e:
        if pythoncom and isinstance(e, pythoncom.com_error):
            print("COM ERROR:", hex(e.hresult), e.excepinfo, file=sys.stderr)
        traceback.print_exc()
        raise

def list_tasks(group: str = DEFAULT_GROUP):
    """指定グループ配下のタスク一覧を取得"""
    require_pywin32()
    svc = connect_service()
    try:
        folder = get_or_create_folder(svc, group_path(group))
    except Exception:
        return []
    return [(t.Name, str(t.NextRunTime), int(t.State)) for t in folder.GetTasks(0)]

def delete_task(task_name, group: str = DEFAULT_GROUP):
    """指定グループ配下のタスクを削除"""
    require_pywin32()
    svc = connect_service()
    folder = get_or_create_folder(svc, group_path(group))
    folder.DeleteTask(task_name, 0)

def run_task_now(task_name, group: str = DEFAULT_GROUP):
    """指定グループ配下のタスクを即時実行"""
    require_pywin32()
    svc = connect_service()
    folder = get_or_create_folder(svc, group_path(group))
    folder.GetTask(task_name).Run("")

def get_task_info(task_name: str, group: str = DEFAULT_GROUP):
    """タスク詳細（URL/スケジュール/開始日/時刻/曜日）を取得"""
    require_pywin32()
    svc = connect_service()
    folder = get_or_create_folder(svc, group_path(group))
    t = folder.GetTask(task_name)
    td = t.Definition

    # Actions から URL を抽出
    try:
        act = td.Actions.Item(1)  # 1-based
    except Exception:
        act = next(iter(td.Actions), None)
    url = extract_url_from_cmdargs(getattr(act, "Arguments", "")) if act else ""

    # Trigger を解析
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

def make_taskdef(svc, runlevel=0):
    """TaskDefinition を作成し、基本設定を適用"""
    td = svc.NewTask(0)
    td.RegistrationInfo.Description = "Created by Python (win32com)"
    td.Principal.LogonType = TASK_LOGON_INTERACTIVE_TOKEN
    td.Principal.RunLevel = runlevel
    s = td.Settings
    s.Enabled = True
    s.StartWhenAvailable = True
    s.AllowDemandStart = True
    s.DisallowStartIfOnBatteries = False
    s.StopIfGoingOnBatteries = False
    s.Hidden = False
    s.ExecutionTimeLimit = "PT0S"  # 実行時間制限なし
    return td

# --- フォーム反映（グローバル関数：App の現在グループを参照） ---------------

def apply_task_to_form(self, task_name: str):
    """選択したタスク内容をフォームへ反映し、全項目をロック"""
    info = get_task_info(task_name, self.group_var.get() or DEFAULT_GROUP)

    # タスク名/URL
    self.name_var.set(info.get("name", ""))
    self.url_var.set(info.get("url", ""))

    # 頻度（表示ラベルへ変換）
    sched_key = info.get("schedule", "ONCE")
    self.schedule_var.set(SCHEDULE_LABELS.get(sched_key, sched_key))

    # 日付/時刻
    self.date_var.set(info.get("start_date", ""))
    self.time_var.set(info.get("start_time", ""))

    # 曜日（WEEKLY のみ）
    weekdays = set(info.get("weekdays", []))
    for code, var in self.weekly_vars.items():
        var.set(code in weekdays)
    if sched_key != "WEEKLY":
        for var in self.weekly_vars.values():
            var.set(False)

    # 全入力ロック＋保存不可
    self._lock_all_inputs()
    self.btn_save.configure(state="disabled")

    self.status.set(f"選択中: {task_name}")

# --- GUI ---------------------------------------------------------------------

class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(APP_TITLE)
        self.geometry("900x700")
        self.minsize(820, 600)

        # ロゴ（任意）
        try:
            if os.path.exists(LOGO_PATH):
                self.logo_img = tk.PhotoImage(file=LOGO_PATH)
                self.iconphoto(True, self.logo_img)
            else:
                self.logo_img = None
        except Exception:
            self.logo_img = None

        # 起動時：default グループを保証
        try:
            ensure_default_group()
        except Exception as e:
            messagebox.showerror("エラー", f"フォルダー初期化に失敗しました: {e}")

        self.create_widgets()
        self.refresh_groups()
        self.refresh_tasks()

    def create_widgets(self):
        pad = {"padx": 8, "pady": 6}

        # --- 作成パネル -------------------------------------------------------
        frm = ttk.LabelFrame(self, text="スケジュール作成")
        frm.pack(fill="x", **pad)

        # グループ選択＋作成
        ttk.Label(frm, text="グループ").grid(row=0, column=0, sticky="w", **pad)
        self.group_var = tk.StringVar(value=DEFAULT_GROUP)
        self.combo_group = ttk.Combobox(frm, textvariable=self.group_var, state="readonly", width=20)
        self.combo_group.grid(row=0, column=1, sticky="w", **pad)
        self.combo_group.bind("<<ComboboxSelected>>", lambda e: self.on_group_changed())

        ttk.Button(frm, text="グループ作成", command=self.on_create_group).grid(row=0, column=2, sticky="w", **pad)

        # タスク名
        ttk.Label(frm, text="タスク名").grid(row=1, column=0, sticky="w", **pad)
        self.name_var = tk.StringVar(value=DEFAULT_TASK_PREFIX)
        self.entry_name = ttk.Entry(frm, textvariable=self.name_var, width=36)
        self.entry_name.grid(row=1, column=1, sticky="w", **pad)

        # URL
        ttk.Label(frm, text="Teamsリンク（URL）").grid(row=2, column=0, sticky="w", **pad)
        self.url_var = tk.StringVar()
        self.entry_url = ttk.Entry(frm, textvariable=self.url_var, width=70)
        self.entry_url.grid(row=2, column=1, columnspan=3, sticky="we", **pad)

        # 頻度
        ttk.Label(frm, text="頻度").grid(row=3, column=0, sticky="w", **pad)
        self.schedule_var = tk.StringVar(value=SCHEDULE_LABELS[DEFAULT_SCHEDULE_KEY])
        self.combo_schedule = ttk.Combobox(frm, textvariable=self.schedule_var, state="readonly",
                                           width=10, values=list(SCHEDULE_LABELS.values()))
        self.combo_schedule.grid(row=3, column=1, sticky="w", **pad)
        self.combo_schedule.bind("<<ComboboxSelected>>", self.on_schedule_changed)

        # 適用開始日
        ttk.Label(frm, text="適用開始日（YYYY-MM-DD）").grid(row=3, column=2, sticky="e", **pad)
        self.date_var = tk.StringVar(value=today_str())
        self.entry_date = ttk.Entry(frm, textvariable=self.date_var, width=14)
        self.entry_date.grid(row=3, column=3, sticky="w", **pad)

        # 時刻
        ttk.Label(frm, text="時刻（HH:MM 24h）").grid(row=4, column=2, sticky="e", **pad)
        self.time_var = tk.StringVar(value=DEFAULT_TIME)
        self.entry_time = ttk.Entry(frm, textvariable=self.time_var, width=10)
        self.entry_time.grid(row=4, column=3, sticky="w", **pad)

        # 曜日チェック（WEEKLY 用）
        self.weekly_vars = {}
        self.days_frame = ttk.Frame(frm)
        self.days_frame.grid(row=4, column=1, sticky="w", **pad)
        for i, (label, code) in enumerate(WEEKDAY_MAP):
            v = tk.BooleanVar(value=(code in DEFAULT_WEEKDAYS))
            self.weekly_vars[code] = v
            ttk.Checkbutton(self.days_frame, text=label, variable=v).grid(row=0, column=i, sticky="w")

        # ボタン
        btn_fr = ttk.Frame(frm); btn_fr.grid(row=5, column=0, columnspan=4, sticky="e", **pad)
        self.btn_save  = ttk.Button(btn_fr, text="作成/更新", command=self.on_create)
        self.btn_clear = ttk.Button(btn_fr, text="クリア", command=self.on_clear_form)
        ttk.Button(btn_fr, text="再読み込み", command=self.refresh_tasks).pack(side="left", padx=4)
        self.btn_save.pack(side="left", padx=4)
        self.btn_clear.pack(side="left", padx=4)

        # --- 一覧パネル -------------------------------------------------------
        list_fr = ttk.LabelFrame(self, text=f"登録済みタスク")
        list_fr.pack(fill="both", expand=True, **pad)

        cols = ("TaskName", "NextRun", "State")
        self.tree = ttk.Treeview(list_fr, columns=cols, show="headings", height=12)
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)
        self.tree.tag_configure("expired", foreground="red")
        for c, t, w in [("TaskName","タスク名",420),("NextRun","次回実行",200),("State","状態",80)]:
            self.tree.heading(c, text=t)
            self.tree.column(c, width=w)
        self.tree.pack(fill="both", expand=True, padx=6, pady=6)

        ctl_fr = ttk.Frame(list_fr); ctl_fr.pack(fill="x", padx=6, pady=6)
        ttk.Button(ctl_fr, text="削除", command=self.on_delete).pack(side="left")
        ttk.Button(ctl_fr, text="修正", command=self.on_enable_edit_fields).pack(side="left", padx=6)

        # 一覧の操作系バインド
        self.tree.bind("<Double-Button-1>", self.on_tree_row_dblrun)  # ダブルクリック＝実行
        self.tree.bind("<Return>", self.on_tree_run_enter)            # Enter＝実行
        self.bind("<Escape>", self.on_escape, add="+")                # Esc＝クリア
        self.AUTO_EDIT_ON_SINGLE_CLICK = True
        self.tree.bind("<<TreeviewSelect>>", self.on_tree_row_select) # 単一クリック＝反映

        # ステータスバー
        self.status = tk.StringVar(value="準備完了")
        ttk.Label(self, textvariable=self.status, anchor="w").pack(fill="x")

    # --- グループ操作 --------------------------------------------------------

    def refresh_groups(self):
        """グループ一覧を読み直し、コンボに反映"""
        try:
            vals = list_groups()
            self.combo_group['values'] = vals
            if self.group_var.get() not in vals:
                self.group_var.set(DEFAULT_GROUP)
        except Exception as e:
            self.status.set(f"グループ読込失敗: {e}")

    def on_create_group(self):
        """グループ（サブフォルダ）新規作成"""
        try:
            name = simpledialog.askstring("グループ作成", "グループ名を入力してください。")
            if not name:
                return
            name = sanitize_name(name)
            if not name or name.lower() == "default":
                messagebox.showwarning("注意", "使用できないグループ名です。")
                return
            require_pywin32()
            svc = connect_service()
            get_or_create_folder(svc, group_path(name))
            self.refresh_groups()
            self.group_var.set(name)
            self.refresh_tasks()
            self.on_clear_form()
            self.status.set(f"グループを作成しました: {name}")
        except Exception as e:
            messagebox.showerror("エラー", f"グループ作成に失敗: {e}")

    def on_group_changed(self):
        """グループ変更時：一覧を切替、フォームは初期化"""
        self.refresh_tasks()
        self.on_clear_form()
        self.status.set(f"グループ切替: {self.group_var.get()}")

    # --- CRUD/UI ハンドラ ----------------------------------------------------

    def on_create(self):
        """作成/更新：現在のグループ配下に登録"""
        try:
            name = sanitize_name(self.name_var.get().strip() or (DEFAULT_TASK_PREFIX + "Task"))

            selected_label = self.schedule_var.get()
            schedule_key = SCHEDULE_FROM_LABEL.get(selected_label, selected_label)

            weekdays = [code for code, v in self.weekly_vars.items() if v.get()] \
                       if schedule_key == "WEEKLY" else []

            current_group = self.group_var.get() or DEFAULT_GROUP
            register_task(
                schedule_key, name, self.url_var.get().strip(),
                self.date_var.get().strip(), self.time_var.get().strip(),
                weekdays, False, current_group
            )
            messagebox.showinfo("完了", f"タスクを作成/更新しました: {group_path(current_group)}\\{name}")
            self.refresh_tasks()
            self.on_clear_form()
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def refresh_tasks(self):
        """一覧を現在グループで更新し、終了済み等は赤色で表示"""
        for i in self.tree.get_children():
            self.tree.delete(i)

        try:
            rows = list_tasks(self.group_var.get() or DEFAULT_GROUP)
        except Exception as e:
            self.status.set(f"一覧取得に失敗: {e}")
            return

        for name, nextrun, state in rows:
            # ▼ 赤色判定：NextRun の形でも簡易チェック
            expired = self._looks_never_runs(nextrun) or self._is_task_inactive(name, state)
            # ▼ タグを付けて挿入（expired → 赤）
            if expired:
                self.tree.insert("", "end", values=(name, nextrun, state), tags=("expired",))
            else:
                self.tree.insert("", "end", values=(name, nextrun, state))

        self.status.set(f"[{self.group_var.get()}] {len(rows)}件のタスク")

    def on_delete(self):
        """選択タスクを削除"""
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("注意", "削除するタスクを選択してください。")
            return
        name = self.tree.item(sel[0], "values")[0]
        if not messagebox.askyesno("削除確認", f"「{name}」を削除してもよろしいですか？"):
            return
        try:
            delete_task(name, self.group_var.get() or DEFAULT_GROUP)
            self.on_clear_form()
            self.refresh_tasks()
            messagebox.showinfo("完了", "削除しました。")
        except Exception as e:
            messagebox.showerror("エラー", str(e))

    def on_enable_edit_fields(self):
        """修正モード：頻度/適用開始日/時刻のみ編集可（曜日は WEEKLY のときのみ）"""
        if not self.name_var.get().strip():
            messagebox.showwarning("注意", "修正するタスクを先に選択してください。")
            return
        # タスク名/URL は常に編集不可
        self.entry_name.configure(state="disabled")
        self.entry_url.configure(state="disabled")
        # 編集可にする項目
        self.combo_schedule.configure(state="readonly")
        self.entry_date.configure(state="normal")
        self.entry_time.configure(state="normal")
        # WEEKLY のときのみ曜日を有効化
        key = SCHEDULE_FROM_LABEL.get(self.schedule_var.get(), self.schedule_var.get())
        self._set_weekday_checks_enabled(key == "WEEKLY")
        # 保存ボタン有効化
        self.btn_save.configure(state="normal")
        self.status.set("修正モード: 頻度/適用開始日/時刻を修正できます")

    # --- 一覧の操作（反映/実行/選択解除） -------------------------------------

    def on_tree_row_select(self, event=None):
        """単一クリックでフォームに反映（AUTO_EDIT_ON_SINGLE_CLICK=True の場合）"""
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

    def on_tree_row_dblrun(self, event):
        """ダブルクリックで即時実行"""
        try:
            region = self.tree.identify("region", event.x, event.y)
            if region not in ("tree", "cell"):
                return
            item_id = self.tree.identify_row(event.y) or (self.tree.selection()[0] if self.tree.selection() else None)
            if not item_id:
                return
            values = self.tree.item(item_id, "values") or ()
            if not values:
                return
            task_name = values[0]
            if not messagebox.askyesno("実行確認", f"「{task_name}」を今すぐ実行しますか？"):
                return
            run_task_now(task_name, self.group_var.get() or DEFAULT_GROUP)
            self.status.set(f"実行要求を送信しました: {task_name}")
        except Exception as e:
            messagebox.showerror("エラー", f"実行に失敗しました: {e}")

    def on_tree_run_enter(self, event=None):
        """Enter キーで即時実行"""
        try:
            item_id = self.tree.focus() or (self.tree.selection()[0] if self.tree.selection() else None)
            if not item_id:
                return
            values = self.tree.item(item_id, "values") or ()
            if not values:
                return
            task_name = values[0]
            run_task_now(task_name, self.group_var.get() or DEFAULT_GROUP)
            self.status.set(f"実行要求を送信しました: {task_name}")
        except Exception as e:
            messagebox.showerror("エラー", f"実行に失敗しました: {e}")

    # --- 入力の有効/無効制御＆クリア ------------------------------------------

    def _set_weekday_checks_enabled(self, enabled: bool):
        """曜日チェック群の有効/無効切り替え"""
        desired = ("!disabled",) if enabled else ("disabled",)
        for w in self.days_frame.winfo_children():
            try:
                w.state(desired)
            except Exception:
                try:
                    w.configure(state=("normal" if enabled else "disabled"))
                except Exception:
                    pass

    def _lock_all_inputs(self):
        """全入力をロック（読み取り専用）"""
        self.entry_name.configure(state="disabled")
        self.entry_url.configure(state="disabled")
        self.entry_date.configure(state="disabled")
        self.entry_time.configure(state="disabled")
        self.combo_schedule.configure(state="disabled")
        self.btn_save.configure(state="disabled")
        self._set_weekday_checks_enabled(False)

    def _unlock_all_inputs(self):
        """編集可能状態に戻す（曜日は頻度に応じて）"""
        self.entry_name.configure(state="normal")
        self.entry_url.configure(state="normal")
        self.entry_date.configure(state="normal")
        self.entry_time.configure(state="normal")
        self.combo_schedule.configure(state="readonly")
        key = SCHEDULE_FROM_LABEL.get(self.schedule_var.get(), self.schedule_var.get())
        self._set_weekday_checks_enabled(key == "WEEKLY")

    def on_schedule_changed(self, event=None):
        """頻度選択変更時：曜日の有効/無効を切替"""
        key = SCHEDULE_FROM_LABEL.get(self.schedule_var.get(), self.schedule_var.get())
        self._set_weekday_checks_enabled(key == "WEEKLY")

    def on_clear_form(self):
        """フォーム全体を初期値に戻し、編集可能にする + 一覧の選択解除"""
        # 編集可能へ
        self._unlock_all_inputs()
        # 選択解除
        self.tree.selection_set(())
        self.tree.focus("")
        # 値を既定へ
        self.name_var.set(DEFAULT_TASK_PREFIX)
        self.url_var.set("")
        self.schedule_var.set(SCHEDULE_LABELS.get(DEFAULT_SCHEDULE_KEY, DEFAULT_SCHEDULE_KEY))
        self.date_var.set(today_str())
        self.time_var.set(DEFAULT_TIME)
        for code, var in self.weekly_vars.items():
            var.set(code in DEFAULT_WEEKDAYS)
        self._set_weekday_checks_enabled(True)  # 既定は WEEKLY

        # 保存ボタンは有効
        self.btn_save.configure(state="normal")
        self.status.set("フォームを初期化しました")

    def on_escape(self, event=None):
        """Esc キーでクリア（コンボドロップダウンが開いている場合はそちらを優先）"""
        try:
            w = self.focus_get()
            if isinstance(w, ttk.Combobox):
                return
        except Exception:
            pass
        self.on_clear_form()

    def _looks_never_runs(self, nextrun_str: str) -> bool:
        """次回実行時刻が無い/無効っぽい場合の簡易判定"""
        if not nextrun_str:
            return True
        s = str(nextrun_str).strip().lower()
        # 既知パターン（環境差に備えて緩めに判定）
        return (
            s in ("none", "0001-01-01 00:00:00")  # 初期値っぽい
            or s.startswith("0001-01-01")         # 0001 年など
            or s == ""                            # 空文字
        )

    def _is_task_inactive(self, name: str, state: int) -> bool:
        """
        行に赤色を付けるべきかを判定。
        基準：
        - Task.State が「無効」（1）なら赤
        - ONCE で開始日時が現在より過去なら赤（＝もう走らない想定）
        - 次回実行時刻が「ほぼ無い」場合も赤
        """
        try:
            if state == 1:  # TASK_STATE_DISABLED
                return True

            info = get_task_info(name, self.group_var.get() or DEFAULT_GROUP)
            # 次回実行が無さそうなら赤
            # （スケジューラの仕様差に備えて念のためチェック）
            # list_tasks の NextRun と info は少し粒度が違うが、併用で堅牢化
            # → NextRun は refresh_tasks で別に見ているので、ここでは補助的に扱う
            if info.get("schedule") == "ONCE":
                d = info.get("start_date") or ""
                t = info.get("start_time") or "00:00"
                if d:
                    try:
                        dt = datetime.strptime(f"{d} {t}", "%Y-%m-%d %H:%M")
                        if dt < datetime.now():
                            return True
                    except Exception:
                        pass
            return False
        except Exception:
            # 取得に失敗したら赤にはしない（安全側）
            return False



# --- メイン ------------------------------------------------------------------

if __name__ == "__main__":
    if os.name != "nt":
        print("Windows only.")
        sys.exit(1)

    # 二重起動防止
    ensure_single_instance()

    # アプリ起動
    app = App()
    app.mainloop()
