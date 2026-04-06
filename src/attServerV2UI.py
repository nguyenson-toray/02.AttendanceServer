# -*- coding: utf-8 -*-
"""
attServerV2UI.py — Attendance Server V2 + tkinter UI
"""
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import queue
import json
import os
import sys
import schedule
import time as time_module
from datetime import datetime, timedelta
from zk import ZK

# ── Import server logic ───────────────────────────────────────────────────────
CWD = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, CWD)
import functions as server

# ── Intercept log writes → UI queue ──────────────────────────────────────────
_log_queue: queue.Queue = queue.Queue()
_orig_write_log = server._write_log


def _patched_write_log(level: str, message: str) -> None:
    """Write to file (original behavior) AND push to UI queue."""
    _orig_write_log(level, message)
    now = datetime.now().strftime('%d-%m-%Y %H:%M:%S')
    _log_queue.put((level.upper(), f'{now} {message}'))


server._write_log = _patched_write_log


# ── Server runner ─────────────────────────────────────────────────────────────
class ServerRunner:
    def __init__(self):
        self._stop_event = threading.Event()
        self._threads: list = []

    def is_running(self) -> bool:
        return any(t.is_alive() for t in self._threads)

    def start(self) -> None:
        self._stop_event.clear()
        self._threads.clear()
        schedule.clear()

        server.read_config()
        server.update_excel_to_mongoDb()

        machine_no = 1
        for ip in server.ip_att_machines:
            machine = ZK(ip, port=4370, timeout=5, password=0, force_udp=False, ommit_ping=False)
            t = threading.Thread(
                target=self._att_thread,
                args=(machine, machine_no),
                name=f'att-{ip}', daemon=True
            )
            self._threads.append(t)
            t.start()
            machine_no += 1

        server.ot_register_detect_qr_and_save()
        server.setup_schedule()

        t_sched = threading.Thread(
            target=self._schedule_thread, daemon=True, name='scheduler')
        self._threads.append(t_sched)
        t_sched.start()

    def stop(self) -> None:
        self._stop_event.set()
        schedule.clear()

    def _att_thread(self, machine: ZK, machineNo: int) -> None:
        server.get_att_log_one_time(machine, machineNo)
        if server.REAL_TIME:
            # live_capture has its own reconnect loop; daemon thread dies on app exit
            server.live_capture_attendance(machine, machineNo)
        else:
            while not self._stop_event.is_set():
                # wait() returns immediately when stop_event is set
                interval = server.schedule_config.get(
                    'att_log_interval_minutes', server.ATT_LOG_INTERVAL_MINUTES)
                self._stop_event.wait(timeout=interval * 60)
                if not self._stop_event.is_set():
                    server.get_att_log_one_time(machine, machineNo)

    def _schedule_thread(self) -> None:
        while not self._stop_event.is_set():
            schedule.run_pending()
            time_module.sleep(1)


# ── Main UI ───────────────────────────────────────────────────────────────────
class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title('Attendance Server V2')
        self.geometry('1150x740')
        self.minsize(900, 600)
        self._runner = ServerRunner()
        self._log_lines: list = []   # (level, text) — stored for filter redraw
        self._last_prune_date = datetime.now().date()
        self._build_ui()
        self._load_config()
        self._poll_log()

    # ── Build UI ──────────────────────────────────────────────────────────────
    def _build_ui(self):
        # Top status bar
        top = tk.Frame(self, bg='#2b2b2b', pady=5)
        top.pack(fill='x')
        tk.Label(top, text='Attendance Server V2', fg='white', bg='#2b2b2b',
                 font=('Segoe UI', 11, 'bold')).pack(side='left', padx=12)
        self._dot = tk.Label(top, text='●', fg='#e74c3c', bg='#2b2b2b',
                             font=('Segoe UI', 14))
        self._dot.pack(side='right', padx=4)
        self._lbl_status = tk.Label(top, text='Stopped', fg='#e74c3c', bg='#2b2b2b',
                                    font=('Segoe UI', 10))
        self._lbl_status.pack(side='right')

        # PanedWindow: left config | right log
        paned = tk.PanedWindow(self, orient='horizontal',
                               sashwidth=5, sashrelief='flat', bg='#cccccc')
        paned.pack(fill='both', expand=True, padx=6, pady=(4, 0))
        paned.add(self._make_config_panel(paned), minsize=340)
        paned.add(self._make_log_panel(paned),    minsize=420)

        # Bottom: Start / Stop
        bottom = tk.Frame(self, pady=8)
        bottom.pack(fill='x')
        self._btn_start = tk.Button(
            bottom, text='▶  START', width=18, font=('Segoe UI', 10, 'bold'),
            bg='#27ae60', fg='white', activebackground='#2ecc71',
            relief='flat', cursor='hand2', command=self._on_start)
        self._btn_start.pack(side='left', padx=24)
        self._btn_stop = tk.Button(
            bottom, text='■  STOP', width=18, font=('Segoe UI', 10, 'bold'),
            bg='#c0392b', fg='white', activebackground='#e74c3c',
            relief='flat', cursor='hand2', state='disabled', command=self._on_stop)
        self._btn_stop.pack(side='left', padx=4)

    # ── Config panel (left, scrollable) ──────────────────────────────────────
    def _make_config_panel(self, parent) -> tk.Widget:
        outer = tk.Frame(parent)
        canvas = tk.Canvas(outer, borderwidth=0, highlightthickness=0)
        vsb = ttk.Scrollbar(outer, orient='vertical', command=canvas.yview)
        canvas.configure(yscrollcommand=vsb.set)
        vsb.pack(side='right', fill='y')
        canvas.pack(side='left', fill='both', expand=True)

        inner = tk.Frame(canvas, padx=8, pady=6)
        win_id = canvas.create_window((0, 0), window=inner, anchor='nw')
        inner.bind('<Configure>',
                   lambda e: canvas.configure(scrollregion=canvas.bbox('all')))
        canvas.bind('<Configure>',
                    lambda e: canvas.itemconfig(win_id, width=e.width))
        canvas.bind_all('<MouseWheel>',
                        lambda e: canvas.yview_scroll(-1 if e.delta > 0 else 1, 'units'))

        self._build_config_fields(inner)
        return outer

    def _build_config_fields(self, p):
        def section(label):
            f = tk.LabelFrame(p, text=label, font=('Segoe UI', 9, 'bold'),
                              fg='#2c3e50', padx=6, pady=4)
            f.pack(fill='x', pady=(0, 6))
            return f

        def field(parent, label, var, browse_cmd=None):
            row = tk.Frame(parent)
            row.pack(fill='x', pady=2)
            tk.Label(row, text=label, width=13, anchor='w',
                     font=('Segoe UI', 9)).pack(side='left')
            tk.Entry(row, textvariable=var, font=('Segoe UI', 9),
                     relief='solid', bd=1).pack(side='left', fill='x', expand=True)
            if browse_cmd:
                tk.Button(row, text='📁', command=browse_cmd, relief='flat',
                          cursor='hand2', font=('Segoe UI', 9),
                          padx=2).pack(side='left', padx=(2, 0))

        # Att Machines
        f = section('Att Machines  (mỗi dòng 1 IP)')
        self._txt_machines = tk.Text(f, height=5, font=('Consolas', 9),
                                     relief='solid', bd=1)
        self._txt_machines.pack(fill='x')

        # File Paths
        f = section('File Paths')
        self._var_aio       = tk.StringVar()
        self._var_resign    = tk.StringVar()
        self._var_maternity = tk.StringVar()
        self._var_ot_folder = tk.StringVar()
        field(f, 'AIO:',       self._var_aio,
              lambda: self._browse_file(self._var_aio))
        field(f, 'Resign:',    self._var_resign,
              lambda: self._browse_file(self._var_resign))
        field(f, 'Maternity:', self._var_maternity,
              lambda: self._browse_file(self._var_maternity))
        field(f, 'OT Folder:', self._var_ot_folder,
              lambda: self._browse_dir(self._var_ot_folder))

        # Sheet Names
        f = section('Sheet Names')
        self._var_sh_aio      = tk.StringVar()
        self._var_sh_resign   = tk.StringVar()
        self._var_sh_leave    = tk.StringVar()
        self._var_sh_pregnant = tk.StringVar()
        self._var_sh_child    = tk.StringVar()
        field(f, 'AIO:',        self._var_sh_aio)
        field(f, 'Resign:',     self._var_sh_resign)
        field(f, 'Mat. leave:', self._var_sh_leave)
        field(f, 'Pregnant:',   self._var_sh_pregnant)
        field(f, 'Child:',      self._var_sh_child)

        # Schedule
        f = section('Schedule')

        # sync_time row (day + time on same line)
        row_sync = tk.Frame(f)
        row_sync.pack(fill='x', pady=2)
        tk.Label(row_sync, text='Sync time:', width=13, anchor='w',
                 font=('Segoe UI', 9)).pack(side='left')
        self._var_sync_day = tk.StringVar()
        day_cb = ttk.Combobox(row_sync, textvariable=self._var_sync_day, width=10,
                              values=['monday','tuesday','wednesday','thursday',
                                      'friday','saturday','sunday'],
                              font=('Segoe UI', 9), state='readonly')
        day_cb.pack(side='left', padx=(0, 4))
        self._var_sync_at = tk.StringVar()
        tk.Entry(row_sync, textvariable=self._var_sync_at, width=7,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left')
        tk.Label(row_sync, text='(HH:MM)', fg='gray',
                 font=('Segoe UI', 8)).pack(side='left', padx=4)

        # excel sync times
        row_ex = tk.Frame(f)
        row_ex.pack(fill='x', pady=2)
        tk.Label(row_ex, text='Excel sync:', width=13, anchor='w',
                 font=('Segoe UI', 9)).pack(side='left')
        self._var_excel_times = tk.StringVar()
        tk.Entry(row_ex, textvariable=self._var_excel_times,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left', fill='x', expand=True)
        tk.Label(f, text='(cách nhau bởi dấu phẩy, vd: 07:00, 09:00, 11:50)',
                 fg='gray', font=('Segoe UI', 8)).pack(anchor='w', padx=2)

        # intervals
        row_int = tk.Frame(f)
        row_int.pack(fill='x', pady=2)
        tk.Label(row_int, text='OT scan:', width=13, anchor='w',
                 font=('Segoe UI', 9)).pack(side='left')
        self._var_ot_interval = tk.StringVar()
        tk.Entry(row_int, textvariable=self._var_ot_interval, width=5,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left')
        tk.Label(row_int, text='min', fg='gray',
                 font=('Segoe UI', 9)).pack(side='left', padx=(2, 16))
        tk.Label(row_int, text='Att log:', font=('Segoe UI', 9)).pack(side='left')
        self._var_att_interval = tk.StringVar()
        tk.Entry(row_int, textvariable=self._var_att_interval, width=5,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left')
        tk.Label(row_int, text='min', fg='gray',
                 font=('Segoe UI', 9)).pack(side='left', padx=2)

        # update window row
        row_win = tk.Frame(f)
        row_win.pack(fill='x', pady=2)
        tk.Label(row_win, text='Update time:', width=13, anchor='w',
                 font=('Segoe UI', 9)).pack(side='left')
        self._var_update_from = tk.StringVar()
        tk.Entry(row_win, textvariable=self._var_update_from, width=7,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left')
        tk.Label(row_win, text='→', fg='gray',
                 font=('Segoe UI', 9)).pack(side='left', padx=4)
        self._var_update_to = tk.StringVar()
        tk.Entry(row_win, textvariable=self._var_update_to, width=7,
                 font=('Segoe UI', 9), relief='solid', bd=1).pack(side='left')
        tk.Label(row_win, text='(HH:MM)', fg='gray',
                 font=('Segoe UI', 8)).pack(side='left', padx=4)

        # Bypass Names
        f = section('Bypass Names  (mỗi dòng 1 tên)')
        self._txt_bypass = tk.Text(f, height=4, font=('Consolas', 9),
                                   relief='solid', bd=1)
        self._txt_bypass.pack(fill='x')

        # Poppler
        f = section('Poppler Path')
        self._var_poppler = tk.StringVar()
        field(f, 'Path:', self._var_poppler,
              lambda: self._browse_dir(self._var_poppler))

        # Buttons
        bf = tk.Frame(p)
        bf.pack(fill='x', pady=4)
        tk.Button(bf, text='💾  Save Config', command=self._save_config,
                  bg='#2980b9', fg='white', activebackground='#3498db',
                  relief='flat', cursor='hand2',
                  font=('Segoe UI', 9)).pack(side='left', padx=(0, 6))
        tk.Button(bf, text='↺  Reload', command=self._load_config,
                  relief='flat', cursor='hand2',
                  font=('Segoe UI', 9)).pack(side='left')

    # ── Log panel (right) ─────────────────────────────────────────────────────
    def _make_log_panel(self, parent) -> tk.Widget:
        outer = tk.Frame(parent)

        # Filter bar
        bar = tk.Frame(outer, pady=4)
        bar.pack(fill='x', padx=4)
        tk.Label(bar, text='Filter:', font=('Segoe UI', 9)).pack(side='left', padx=(0, 4))
        self._filter_var = tk.StringVar(value='ALL')
        for val, color in (('ALL', '#d4d4d4'), ('INFO', '#9cdcfe'),
                           ('ERROR', '#f48771'), ('DEBUG', '#808080')):
            tk.Radiobutton(bar, text=val, variable=self._filter_var, value=val,
                           fg=color, font=('Segoe UI', 9),
                           command=self._apply_filter).pack(side='left', padx=2)
        tk.Button(bar, text='Clear', command=self._clear_log,
                  relief='flat', cursor='hand2',
                  font=('Segoe UI', 9)).pack(side='right', padx=4)

        # Text area + scrollbars
        txt_frame = tk.Frame(outer)
        txt_frame.pack(fill='both', expand=True, padx=4, pady=(0, 4))

        ys = ttk.Scrollbar(txt_frame, orient='vertical')
        xs = ttk.Scrollbar(txt_frame, orient='horizontal')
        self._log_text = tk.Text(
            txt_frame, state='disabled', wrap='none',
            font=('Consolas', 9), bg='#1e1e1e', fg='#d4d4d4',
            relief='flat', selectbackground='#264f78',
            yscrollcommand=ys.set, xscrollcommand=xs.set)
        ys.configure(command=self._log_text.yview)
        xs.configure(command=self._log_text.xview)
        ys.pack(side='right', fill='y')
        xs.pack(side='bottom', fill='x')
        self._log_text.pack(fill='both', expand=True)

        # Color tags per log level
        self._log_text.tag_config('INFO',  foreground='#9cdcfe')
        self._log_text.tag_config('ERROR', foreground='#f48771')
        self._log_text.tag_config('DEBUG', foreground='#808080')

        return outer

    # ── Config load / save ────────────────────────────────────────────────────
    def _load_config(self):
        try:
            with open(server.CONFIG_PATH, 'r', encoding='utf-8') as f:
                cfg = json.load(f)
            self._txt_machines.delete('1.0', 'end')
            self._txt_machines.insert('1.0', '\n'.join(cfg.get('att_machines', [])))
            paths = cfg.get('paths', {})
            self._var_aio.set(paths.get('aio', ''))
            self._var_resign.set(paths.get('resign', ''))
            self._var_maternity.set(paths.get('maternity', ''))
            self._var_ot_folder.set(paths.get('ot_folder', ''))
            sheets = cfg.get('sheets', {})
            self._var_sh_aio.set(sheets.get('aio', ''))
            self._var_sh_resign.set(sheets.get('resign', ''))
            self._var_sh_leave.set(sheets.get('maternity_leave', ''))
            self._var_sh_pregnant.set(sheets.get('maternity_pregnant', ''))
            self._var_sh_child.set(sheets.get('maternity_child', ''))
            sc = cfg.get('schedule', {})
            self._var_sync_day.set(sc.get('sync_time_day', 'sunday'))
            self._var_sync_at.set(sc.get('sync_time_at', '06:00'))
            self._var_excel_times.set(', '.join(sc.get('excel_sync_times', [])))
            self._var_ot_interval.set(str(sc.get('ot_scan_interval_minutes', 10)))
            self._var_att_interval.set(str(sc.get('att_log_interval_minutes', 6)))
            self._var_update_from.set(sc.get('update_time_from', '17:00'))
            self._var_update_to.set(sc.get('update_time_to', '22:00'))
            self._txt_bypass.delete('1.0', 'end')
            self._txt_bypass.insert('1.0', '\n'.join(cfg.get('bypass_names', [])))
            self._var_poppler.set(cfg.get('poppler_path', ''))
        except FileNotFoundError:
            messagebox.showwarning('Config', f'config.json not found:\n{server.CONFIG_PATH}')
        except Exception as e:
            messagebox.showerror('Load Config Error', str(e))

    def _save_config(self):
        try:
            machines = [ip.strip()
                        for ip in self._txt_machines.get('1.0', 'end').splitlines()
                        if ip.strip()]
            bypass = [n.strip()
                      for n in self._txt_bypass.get('1.0', 'end').splitlines()
                      if n.strip()]
            cfg = {
                'att_machines': machines,
                'poppler_path': self._var_poppler.get().strip(),
                'bypass_names': bypass,
                'paths': {
                    'aio':       self._var_aio.get().strip(),
                    'resign':    self._var_resign.get().strip(),
                    'maternity': self._var_maternity.get().strip(),
                    'ot_folder': self._var_ot_folder.get().strip(),
                },
                'sheets': {
                    'aio':                self._var_sh_aio.get().strip(),
                    'resign':             self._var_sh_resign.get().strip(),
                    'maternity_leave':    self._var_sh_leave.get().strip(),
                    'maternity_pregnant': self._var_sh_pregnant.get().strip(),
                    'maternity_child':    self._var_sh_child.get().strip(),
                },
                'schedule': {
                    'sync_time_day':           self._var_sync_day.get().strip(),
                    'sync_time_at':            self._var_sync_at.get().strip(),
                    'excel_sync_times':        [t.strip() for t in
                                                self._var_excel_times.get().split(',')
                                                if t.strip()],
                    'ot_scan_interval_minutes': int(self._var_ot_interval.get().strip() or 10),
                    'att_log_interval_minutes': int(self._var_att_interval.get().strip() or 6),
                    'update_time_from': self._var_update_from.get().strip() or '17:00',
                    'update_time_to':   self._var_update_to.get().strip()   or '22:00',
                },
            }
            with open(server.CONFIG_PATH, 'w', encoding='utf-8') as f:
                json.dump(cfg, f, ensure_ascii=False, indent=2)
            messagebox.showinfo('Saved', 'Config saved.')
        except Exception as e:
            messagebox.showerror('Save Config Error', str(e))

    # ── Browse helpers ────────────────────────────────────────────────────────
    def _browse_file(self, var: tk.StringVar):
        path = filedialog.askopenfilename(
            filetypes=[('Excel files', '*.xlsx *.xls'), ('All files', '*.*')])
        if path:
            var.set(path)

    def _browse_dir(self, var: tk.StringVar):
        path = filedialog.askdirectory()
        if path:
            var.set(path)

    # ── Log ───────────────────────────────────────────────────────────────────
    def _poll_log(self):
        """Drain queue every 200 ms on main thread → append to Text widget."""
        try:
            while True:
                level, text = _log_queue.get_nowait()
                self._log_lines.append((level, text))
                self._append_line(level, text)
        except queue.Empty:
            pass
        today = datetime.now().date()
        if today != self._last_prune_date:
            self._prune_old_logs()
            self._last_prune_date = today
        self.after(200, self._poll_log)

    def _prune_old_logs(self):
        """Remove UI entries older than yesterday. Does NOT touch log files on disk."""
        cutoff = datetime.now().date() - timedelta(days=1)
        kept = []
        for level, text in self._log_lines:
            try:
                if datetime.strptime(text[:10], '%d-%m-%Y').date() >= cutoff:
                    kept.append((level, text))
            except ValueError:
                kept.append((level, text))
        if len(kept) == len(self._log_lines):
            return
        self._log_lines = kept
        self._apply_filter()

    def _append_line(self, level: str, text: str):
        if self._filter_var.get() not in ('ALL', level):
            return
        self._log_text.configure(state='normal')
        self._log_text.insert('end', text + '\n', level)
        self._log_text.see('end')
        self._log_text.configure(state='disabled')

    def _apply_filter(self):
        filt = self._filter_var.get()
        self._log_text.configure(state='normal')
        self._log_text.delete('1.0', 'end')
        for level, text in self._log_lines:
            if filt == 'ALL' or level == filt:
                self._log_text.insert('end', text + '\n', level)
        self._log_text.see('end')
        self._log_text.configure(state='disabled')

    def _clear_log(self):
        self._log_lines.clear()
        self._log_text.configure(state='normal')
        self._log_text.delete('1.0', 'end')
        self._log_text.configure(state='disabled')

    # ── Start / Stop ──────────────────────────────────────────────────────────
    def _on_start(self):
        if self._runner.is_running():
            return
        self._btn_start.configure(state='disabled')
        self._btn_stop.configure(state='normal')
        self._set_status(running=True)
        threading.Thread(target=self._start_worker, daemon=True).start()

    def _start_worker(self):
        """Runs in background thread so UI stays responsive during startup."""
        try:
            self._runner.start()
        except Exception as e:
            _log_queue.put(('ERROR',
                f'{datetime.now():%d-%m-%Y %H:%M:%S} [MAIN] ERROR: Start failed: {e}'))
            self.after(0, lambda: self._set_status(running=False))
            self.after(0, lambda: self._btn_start.configure(state='normal'))
            self.after(0, lambda: self._btn_stop.configure(state='disabled'))

    def _on_stop(self):
        self._runner.stop()
        self._btn_start.configure(state='normal')
        self._btn_stop.configure(state='disabled')
        self._set_status(running=False)
        _log_queue.put(('INFO',
            f'{datetime.now():%d-%m-%Y %H:%M:%S} [MAIN] INFO: Server stopped by user.'))

    def _set_status(self, running: bool):
        color = '#27ae60' if running else '#e74c3c'
        text  = 'Running'  if running else 'Stopped'
        self._dot.configure(fg=color)
        self._lbl_status.configure(fg=color, text=text)


if __name__ == '__main__':
    app = App()
    app.mainloop()
