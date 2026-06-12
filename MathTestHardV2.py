"""
Math Test — Version B: Hard (Scratch Work Required)
=====================================================
All questions require the user to write down working steps.
No question should be doable purely in your head, but none
requires a calculator — clean integer or simple fraction answers.

Fixed 35 seconds per question throughout. No tier adaptation —
difficulty stays constant to isolate the scratch-work effect.

Flow:
  1. Intake screen (participant info)
  2. Pre-test relaxation screen (3 min, relax.wav)
  3. "Open eyes" sound plays, then instructions window
  4. Math test (10 questions, 35 s each)
  5. Post-test relaxation screen (3 min, relax.wav)
  6. Ending screen

Question categories:
  • Multi-step linear algebra (two equations, substitution)
  • Quadratic equations (factor by hand)
  • Geometry (area + perimeter combos, similar triangles, Pythagorean)
  • Arithmetic sequences & series (sum of N terms)
  • Percent chain problems (two successive changes)
  • Ratio + proportion word problems (multi-step)
  • Work / rate problems (three workers)
  • Age word problems
  • Digit / number theory problems
  • Mixture problems
"""

import tkinter as tk
from tkinter import ttk
import random
import csv
import time
import threading
import os
import sys
import math
import fractions
import re

# ---------------------------------------------------------------------------
def get_resource_path(relative_path: str) -> str:
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

# ---------------------------------------------------------------------------
TOTAL_QUESTIONS   = 10
COUNTDOWN         = 5
BASE_DATA_PATH    = './data/'
QUESTION_TIME_SEC = 35          # fixed — all questions get 35 s
NICLA_ENABLED     = True
NICLA_CHAR_UUID   = "19b10001-e8f2-537e-4f6c-d104768a1214"
NICLA_DEVICE_NAME = "NiclaSenseME"
TEST_LABEL        = "Hard — Scratch Work"
XLSX_PREFIX       = "mathTest_hard_"
RELAX_DURATION    = 180         # 3 minutes

MAX_PAYOUT        = 50.0        # dollars
PENALTY_PER_WRONG = 3.0         # dollars deducted per wrong or missed question

# ---------------------------------------------------------------------------
# Intake screen styling
# ---------------------------------------------------------------------------
IN_BG       = "#ffffff"
IN_ENTRY_BG = "#f5f5f5"
IN_ACCENT   = "#1e3a5f"
IN_TEXT     = "#111111"
IN_MUTED    = "#666666"
IN_ERROR    = "#cc3333"
IN_BORDER   = "#dddddd"
IN_BTN_FG   = "#ffffff"

# ---------------------------------------------------------------------------
# Audio helpers — stdlib only, no pygame required
# Windows  : winsound  (built-in)
# macOS    : afplay    (built-in)
# Linux    : aplay     (usually pre-installed; falls back to paplay/ffplay)
# ---------------------------------------------------------------------------
import platform as _platform
import subprocess as _subprocess

_SYSTEM = _platform.system()   # "Windows" | "Darwin" | "Linux"

_loop_stop_evt = None
_loop_thread   = None

_active_audio_proc = None
_active_audio_lock = threading.Lock()


def _stop_all_audio():
    """Kill any currently playing audio immediately (all platforms)."""
    global _active_audio_proc
    if _SYSTEM == "Windows":
        try:
            import winsound
            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception:
            pass
    else:
        with _active_audio_lock:
            if _active_audio_proc and _active_audio_proc.poll() is None:
                try:
                    _active_audio_proc.terminate()
                    _active_audio_proc.wait(timeout=1)
                except Exception:
                    pass
            _active_audio_proc = None


def _run_subprocess_audio(cmd):
    """Run one audio subprocess, tracking it so _stop_all_audio can kill it."""
    global _active_audio_proc
    try:
        proc = _subprocess.Popen(cmd,
                                  stdout=_subprocess.DEVNULL,
                                  stderr=_subprocess.DEVNULL)
        with _active_audio_lock:
            _active_audio_proc = proc
        proc.wait()
        with _active_audio_lock:
            if _active_audio_proc is proc:
                _active_audio_proc = None
    except FileNotFoundError:
        pass


def _get_player_cmds(path):
    if _SYSTEM == "Darwin":
        return [["afplay", path]]
    else:
        return [["aplay", "-q", path],
                ["paplay", path],
                ["ffplay", "-nodisp", "-autoexit", "-loglevel", "quiet", path]]


def _play_audio_loop(path, stop_event):
    """
    Loop a .wav file until stop_event is set.
    """
    if _SYSTEM == "Windows":
        try:
            import winsound
            while not stop_event.is_set():
                winsound.PlaySound(path, winsound.SND_FILENAME)
            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception as e:
            print(f"[Audio] loop error: {e}")
    else:
        cmds = _get_player_cmds(path)
        while not stop_event.is_set():
            for cmd in cmds:
                if stop_event.is_set():
                    return
                global _active_audio_proc
                try:
                    proc = _subprocess.Popen(cmd,
                                              stdout=_subprocess.DEVNULL,
                                              stderr=_subprocess.DEVNULL)
                    with _active_audio_lock:
                        _active_audio_proc = proc
                    while proc.poll() is None:
                        if stop_event.is_set():
                            proc.terminate()
                            try:
                                proc.wait(timeout=1)
                            except Exception:
                                pass
                            return
                        time.sleep(0.05)
                    with _active_audio_lock:
                        if _active_audio_proc is proc:
                            _active_audio_proc = None
                    break
                except FileNotFoundError:
                    continue
            else:
                print("[Audio] no suitable player found for loop")
                return


def _play_audio_once(path):
    """Play a .wav file exactly once (blocking). Run in a daemon thread."""
    _stop_all_audio()
    if _SYSTEM == "Windows":
        try:
            import winsound
            winsound.PlaySound(path, winsound.SND_FILENAME)
        except Exception as e:
            print(f"[Audio] once error: {e}")
    else:
        for cmd in _get_player_cmds(path):
            try:
                _run_subprocess_audio(cmd)
                return
            except Exception:
                continue
        print("[Audio] no suitable player found for one-shot")


# ---------------------------------------------------------------------------
# Excel / data helpers
# ---------------------------------------------------------------------------
def get_user_data_path(username):
    """Return the per-user data folder path, creating it if needed."""
    folder = os.path.join(BASE_DATA_PATH, username)
    os.makedirs(folder, exist_ok=True)
    return folder

def get_xlsx_path(username):
    return os.path.join(get_user_data_path(username),
                        XLSX_PREFIX + username + '.xlsx')

def create_xlsx_skeleton(path, participant_data):
    """
    Create the Excel file immediately when the test starts so data is
    preserved even if the session ends early. Trial rows are appended
    live via append_trial_to_xlsx().
    """
    try:
        from openpyxl import Workbook
        from openpyxl.styles import Font, PatternFill, Alignment

        wb = Workbook()
        ws = wb.active
        ws.title = "Results"

        accent = "1E3A5F"
        white  = "FFFFFF"
        header_font  = Font(name="Arial", bold=True, color=white)
        header_fill  = PatternFill("solid", start_color=accent)
        center_align = Alignment(horizontal="center", vertical="center")

        pd = participant_data
        meta_rows = [
            ["Participant", pd.get("username", "")],
            ["First Name",  pd.get("first_name", "")],
            ["Last Name",   pd.get("last_name", "")],
            ["Email",       pd.get("email", "")],
            ["Age",         pd.get("age", "")],
            ["Gender",      pd.get("gender", "")],
            ["Test",        TEST_LABEL],
            ["Start Time",  time.strftime('%m/%d/%Y %H:%M:%S')],
            [],
        ]
        for row in meta_rows:
            ws.append(row)
            if row:
                ws.cell(ws.max_row, 1).font = Font(name="Arial", bold=True)

        col_headers = [
            "Timestamp", "Relative Time (ms)", "Response Time (ms)",
            "Q Time Allowed (s)", "Question",
            "Correct Answer", "Selected Option", "Correctness",
            "HM_Samples", "HM_Mean_Mag", "HM_Max_Mag",
        ]
        ws.append(col_headers)
        hdr_row = ws.max_row
        for col_idx, _ in enumerate(col_headers, 1):
            cell = ws.cell(hdr_row, col_idx)
            cell.font       = header_font
            cell.fill       = header_fill
            cell.alignment  = center_align

        col_widths = [20, 20, 20, 18, 60, 16, 18, 12, 12, 14, 14]
        for i, w in enumerate(col_widths, 1):
            ws.column_dimensions[
                ws.cell(1, i).column_letter].width = w

        # ---- Session Timeline sheet ----
        ts = wb.create_sheet("Session Timeline")
        tl_headers = ["Event", "Timestamp"]
        ts.append(tl_headers)
        tl_hdr_row = ts.max_row
        for col_idx in range(1, 3):
            cell = ts.cell(tl_hdr_row, col_idx)
            cell.font      = header_font
            cell.fill      = header_fill
            cell.alignment = center_align
        ts.column_dimensions["A"].width = 32
        ts.column_dimensions["B"].width = 24

        wb.save(path)
        print("Excel file created:", os.path.abspath(path))
    except Exception as e:
        print(f"[Excel] Failed to create skeleton: {e}")

def append_trial_to_xlsx(path, row_data):
    """Append a single trial row to the existing Excel file."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        wb.active.append(row_data)
        wb.save(path)
    except Exception as e:
        print(f"[Excel] append error: {e}")

def write_summary_to_xlsx(path, summary):
    """Write a summary block below the trial data."""
    try:
        from openpyxl import load_workbook
        from openpyxl.styles import Font, PatternFill

        wb   = load_workbook(path)
        ws   = wb.active
        ws.append([])
        ws.append(["=== SUMMARY ==="])
        ws.cell(ws.max_row, 1).font = Font(name="Arial", bold=True)

        for key, val in summary.items():
            ws.append([key, val])

        wb.save(path)
    except Exception as e:
        print(f"[Excel] summary write error: {e}")


def append_timeline_event(path, event_name):
    """Append a timestamped event row to the Session Timeline sheet."""
    try:
        from openpyxl import load_workbook
        wb = load_workbook(path)
        ts = wb["Session Timeline"]
        ts.append([event_name, time.strftime('%m/%d/%Y %H:%M:%S')])
        wb.save(path)
    except Exception as e:
        print(f"[Excel] timeline append error: {e}")


# ---------------------------------------------------------------------------
# Relaxation Screen
# ---------------------------------------------------------------------------
class RelaxationScreen:
    """
    Calming full-screen relaxation window matching MathTestEasyV2 design.
    Plays relax.wav in a loop; fires OpenEyes.wav at the end.
    """

    BG          = "#0a0e1a"
    RING_OUTER  = "#1a2a4a"
    RING_INNER  = "#162038"
    GLOW_COLOR  = "#4a90d9"
    TEXT_MAIN   = "#c8d8f0"
    TEXT_DIM    = "#4a6080"
    TEXT_BREATH = "#7ab0e0"

    INHALE_MS   = 4000
    HOLD_MS     = 1500
    EXHALE_MS   = 5000
    REST_MS     = 1000

    def __init__(self, master, duration=RELAX_DURATION, on_done=None,
                 xlsx_path=None, period_label="Relaxation",
                 play_open_eyes=True):
        self.master        = master
        self.duration      = duration
        self.on_done       = on_done
        self.xlsx_path     = xlsx_path
        self.period_label  = period_label
        self.play_open_eyes = play_open_eyes
        self.running       = True
        self._stop_evt     = threading.Event()

        self.master.title("Relaxation")
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        self.master.configure(bg=self.BG)

        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        self.master.geometry(f"{sw}x{sh}+0+0")
        try:
            self.master.attributes("-fullscreen", True)
        except Exception:
            pass
        try:
            self.master.state("zoomed")
        except Exception:
            pass

        self.sw = sw
        self.sh = sh

        self._breath_phase  = "inhale"
        self._breath_frac   = 0.0
        self._anim_step_ms  = 40
        self._glow_items    = []

        self._base = min(sw, sh)
        self._cx   = sw // 2
        self._cy   = int(sh * 0.46)

        if self.xlsx_path and os.path.exists(self.xlsx_path):
            threading.Thread(
                target=append_timeline_event,
                args=(self.xlsx_path, f"{self.period_label} Start"),
                daemon=True).start()

        music_path = get_resource_path('relax.wav')
        threading.Thread(
            target=_play_audio_loop,
            args=(music_path, self._stop_evt),
            daemon=True).start()

        self._audio_timer = threading.Timer(
            float(duration), self._on_audio_timeout)
        self._audio_timer.daemon = True
        self._audio_timer.start()

        self._build_ui()

        self.master.bind("<Right>", lambda e: self._skip())

        self._tick_second()
        self._breath_step()

    def _fs(self, base_pt):
        return max(8, int(base_pt * self._base / 1080))

    def _build_ui(self):
        b  = self._base
        sw = self.sw
        sh = self.sh

        self.canvas = tk.Canvas(
            self.master, bg=self.BG,
            highlightthickness=0, bd=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

        cx, cy = self._cx, self._cy
        for r_frac, alpha_col in [
            (0.42, "#111828"), (0.35, "#121a30"),
            (0.28, "#131c34"), (0.21, "#141e38"),
        ]:
            r = int(b * r_frac)
            self.canvas.create_oval(
                cx - r, cy - r, cx + r, cy + r,
                fill=alpha_col, outline="")

        self._min_r = int(b * 0.12)
        self._max_r = int(b * 0.22)
        self._circle_item  = None
        self._glow_items   = []
        self._draw_breath_circle(0.0)

        self.canvas.create_text(
            cx, int(sh * 0.10),
            text="Relax",
            font=(self._font("thin"), self._fs(52)),
            fill=self.TEXT_MAIN,
            anchor="center")

        self.canvas.create_text(
            cx, int(sh * 0.165),
            text="Close your eyes and rest",
            font=(self._font("light"), self._fs(20)),
            fill=self.TEXT_DIM,
            anchor="center")

        pad = int(b * 0.03)
        self._countdown_text = self.canvas.create_text(
            sw - pad, sh - pad,
            text=self._fmt_time(self.duration),
            font=(self._font("mono"), self._fs(15)),
            fill=self.TEXT_DIM,
            anchor="se")

        self.canvas.create_text(
            cx, int(sh * 0.93),
            text="You will be notified when it is time to continue.",
            font=(self._font("light"), self._fs(13)),
            fill=self.TEXT_DIM,
            anchor="center")

    @staticmethod
    def _font(style="regular"):
        fonts = {
            "thin":    ("Helvetica Neue", "Segoe UI Light", "Ubuntu Light",  "Helvetica"),
            "light":   ("Helvetica Neue", "Segoe UI",       "Ubuntu",        "Helvetica"),
            "regular": ("Helvetica Neue", "Segoe UI",       "Ubuntu",        "Arial"),
            "mono":    ("Courier New",    "Consolas",        "DejaVu Sans Mono", "Courier"),
        }
        return fonts.get(style, ("Arial",))[0]

    @staticmethod
    def _fmt_time(secs):
        m, s = divmod(secs, 60)
        return f"{m}:{s:02d}"

    def _lerp_color(self, frac):
        r0, g0, b0 = 0x1a, 0x3a, 0x6a
        r1, g1, b1 = 0x50, 0xb8, 0xe8
        r = int(r0 + (r1 - r0) * frac)
        g = int(g0 + (g1 - g0) * frac)
        b = int(b0 + (b1 - b0) * frac)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw_breath_circle(self, frac):
        cx, cy = self._cx, self._cy
        r      = self._min_r + int((self._max_r - self._min_r) * frac)
        color  = self._lerp_color(frac)

        for item in self._glow_items:
            self.canvas.delete(item)
        self._glow_items = []
        if self._circle_item:
            self.canvas.delete(self._circle_item)

        glow_layers = [
            (r + int(self._base * 0.070), "#0d1525"),
            (r + int(self._base * 0.050), "#0f1a30"),
            (r + int(self._base * 0.033), "#122040"),
            (r + int(self._base * 0.018), "#1a3058"),
            (r + int(self._base * 0.008), "#2a5080"),
        ]
        for gr, gc in glow_layers:
            item = self.canvas.create_oval(
                cx - gr, cy - gr, cx + gr, cy + gr,
                fill=gc, outline="")
            self._glow_items.append(item)

        self._circle_item = self.canvas.create_oval(
            cx - r, cy - r, cx + r, cy + r,
            fill=color, outline="")

        hr = max(3, int(r * 0.18))
        hx = cx - int(r * 0.38)
        hy = cy - int(r * 0.38)
        self._glow_items.append(
            self.canvas.create_oval(
                hx - hr, hy - hr, hx + hr, hy + hr,
                fill="#a8d8f8", outline=""))

    def _breath_step(self):
        if not self.running:
            return

        phase = self._breath_phase

        if phase == "inhale":
            self._breath_frac += self._anim_step_ms / self.INHALE_MS
            if self._breath_frac >= 1.0:
                self._breath_frac = 1.0
                self._breath_phase = "hold_in"
        elif phase == "hold_in":
            self._breath_phase = "exhale"
            try:
                self._draw_breath_circle(self._breath_frac)
            except tk.TclError:
                return
            self.master.after(self.HOLD_MS, self._breath_step)
            return
        elif phase == "exhale":
            self._breath_frac -= self._anim_step_ms / self.EXHALE_MS
            if self._breath_frac <= 0.0:
                self._breath_frac = 0.0
                self._breath_phase = "rest"
        else:
            self._breath_phase = "inhale"
            try:
                self._draw_breath_circle(self._breath_frac)
            except tk.TclError:
                return
            self.master.after(self.REST_MS, self._breath_step)
            return

        import math as _math
        eased = (1 - _math.cos(self._breath_frac * _math.pi)) / 2

        try:
            self._draw_breath_circle(eased)
        except tk.TclError:
            return

        self.master.after(self._anim_step_ms, self._breath_step)

    def _tick_second(self):
        if not self.running:
            return
        if self.duration > 0:
            try:
                self.canvas.itemconfig(
                    self._countdown_text,
                    text=self._fmt_time(self.duration))
                self.duration -= 1
                try:
                    self.master.after(1000, self._tick_second)
                except tk.TclError:
                    pass
            except tk.TclError:
                pass
        else:
            self._finish()

    def _on_audio_timeout(self):
        """
        Called by threading.Timer after exactly RELAX_DURATION seconds.
        Stops the relax loop and optionally plays OpenEyes.wav.
        """
        self._stop_evt.set()
        if self.play_open_eyes:
            eyes_path = get_resource_path('OpenEyes.wav')
            threading.Thread(target=_play_audio_once, args=(eyes_path,),
                             daemon=True).start()

    def _skip(self):
        if hasattr(self, '_audio_timer'):
            self._audio_timer.cancel()
        self._stop_evt.set()
        self._finish()

    def _finish(self):
        if not self.running:
            return
        self.running = False
        self._stop_evt.set()
        if hasattr(self, '_audio_timer'):
            self._audio_timer.cancel()
        if self.on_done:
            try:
                self.master.after(50, self._do_finish)
            except tk.TclError:
                self._do_finish()
        else:
            self._do_finish()

    def _do_finish(self):
        cb = self.on_done
        self.on_done = None
        if self.xlsx_path and os.path.exists(self.xlsx_path):
            append_timeline_event(self.xlsx_path, f"{self.period_label} End")
        try:
            self.master.destroy()
        except tk.TclError:
            pass
        if cb:
            cb()

    def on_close(self):
        self.running = False
        self._stop_evt.set()
        if hasattr(self, '_audio_timer'):
            self._audio_timer.cancel()
        self.on_done = None
        time.sleep(0.05)
        try:
            self.master.destroy()
        except tk.TclError:
            pass


# ---------------------------------------------------------------------------
# Instructions Window  (shown after pre-test relaxation)
# ---------------------------------------------------------------------------
class InstructionsWindow:
    """
    Plays the 'open eyes' sound (already triggered by RelaxationScreen),
    then shows exam-specific instructions before the hard math test begins.
    """

    BG        = "#0a0e1a"
    CARD_BG   = "#111828"
    BORDER    = "#1e3050"
    TEXT_HEAD = "#c8d8f0"
    TEXT_BODY = "#8aa8cc"
    TEXT_DIM  = "#4a6080"
    ACCENT    = "#4a90d9"
    BTN_BG    = "#1e5090"
    BTN_HOV   = "#2870c0"

    def __init__(self, master, on_ready=None):
        self.master   = master
        self.on_ready = on_ready

        self.master.title("Test Instructions")
        self.master.configure(bg=self.BG)
        self.master.protocol("WM_DELETE_WINDOW", self._begin)

        sw = self.master.winfo_screenwidth()
        sh = self.master.winfo_screenheight()
        self._base = min(sw, sh)

        ww = max(520, min(int(sw * 0.52), 820))
        wh = max(520, min(int(sh * 0.75), 760))
        x  = (sw - ww) // 2
        y  = (sh - wh) // 2
        self.master.geometry(f"{ww}x{wh}+{x}+{y}")
        self.master.resizable(True, True)

        self._ww = ww
        self._wh = wh

        self._build_ui()

    def _fs(self, pt):
        return max(8, int(pt * self._base / 1080))

    def _build_ui(self):
        fs = self._fs

        outer = tk.Frame(self.master, bg=self.BG)
        outer.pack(fill=tk.BOTH, expand=True)

        card_wrap = tk.Frame(outer, bg=self.BORDER)
        card_wrap.place(relx=0.5, rely=0.5, anchor="center",
                        relwidth=0.88, relheight=0.92)

        card = tk.Frame(card_wrap, bg=self.CARD_BG)
        card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill=tk.BOTH, expand=True,
                   padx=int(self._ww * 0.07),
                   pady=int(self._wh * 0.05))

        # Header
        tk.Label(
            inner,
            text="Time to Wake Up",
            font=("Helvetica Neue", fs(26), "bold"),
            bg=self.CARD_BG, fg=self.TEXT_HEAD
        ).pack(anchor="w")

        tk.Label(
            inner,
            text="Your hard math test is about to begin.",
            font=("Helvetica Neue", fs(13)),
            bg=self.CARD_BG, fg=self.TEXT_DIM
        ).pack(anchor="w", pady=(2, 0))

        # Divider
        tk.Frame(inner, bg=self.BORDER, height=1).pack(
            fill=tk.X, pady=(int(self._wh * 0.025), int(self._wh * 0.030)))

        # Instruction bullets
        bullets = [
            ("10 questions total",
             "This test has exactly 10 questions. Answer each one carefully."),
            ("35 seconds per question",
             "Each question has a 35-second timer. Questions advance automatically when time runs out."),
            ("Scratch work required",
             "These questions need working-out steps. Write your work on paper, then select the closest answer."),
            ("Maximum payout: $50",
             f"You earn ${MAX_PAYOUT:.0f} if you answer all 10 questions correctly. "
             f"Each wrong answer or missed question deducts ${PENALTY_PER_WRONG:.0f} from your total."),
            ("Four choices per question",
             "Each question shows options A, B, C, and D. Select the correct one."),
        ]

        for title, body in bullets:
            row = tk.Frame(inner, bg=self.CARD_BG)
            row.pack(fill=tk.X, pady=(0, int(self._wh * 0.016)))

            dot = tk.Frame(row, bg=self.ACCENT,
                           width=int(self._base * 0.006),
                           height=int(self._base * 0.006))
            dot.pack(side=tk.LEFT, anchor="n",
                     padx=(0, int(self._ww * 0.025)),
                     pady=(int(fs(14) * 0.35), 0))
            dot.pack_propagate(False)

            text_col = tk.Frame(row, bg=self.CARD_BG)
            text_col.pack(side=tk.LEFT, fill=tk.X, expand=True)

            tk.Label(
                text_col, text=title,
                font=("Helvetica Neue", fs(14), "bold"),
                bg=self.CARD_BG, fg=self.TEXT_HEAD,
                anchor="w", justify="left"
            ).pack(anchor="w")

            tk.Label(
                text_col, text=body,
                font=("Helvetica Neue", fs(12)),
                bg=self.CARD_BG, fg=self.TEXT_BODY,
                anchor="w", justify="left",
                wraplength=int(self._ww * 0.65)
            ).pack(anchor="w")

        tk.Frame(inner, bg=self.CARD_BG).pack(fill=tk.BOTH, expand=True)

        self._btn = tk.Button(
            inner,
            text="I'm Ready — Begin Test →",
            font=("Helvetica Neue", fs(14), "bold"),
            bg=self.BTN_BG, fg="white",
            activebackground=self.BTN_HOV, activeforeground="white",
            relief=tk.FLAT, bd=0,
            padx=int(self._ww * 0.04), pady=int(self._wh * 0.022),
            cursor="hand2",
            command=self._begin
        )
        self._btn.pack(fill=tk.X)
        self._btn.bind("<Enter>",
                       lambda e: self._btn.config(bg=self.BTN_HOV))
        self._btn.bind("<Leave>",
                       lambda e: self._btn.config(bg=self.BTN_BG))

    def _begin(self):
        cb = self.on_ready
        self.on_ready = None
        try:
            self.master.destroy()
        except tk.TclError:
            pass
        if cb:
            cb()


# ---------------------------------------------------------------------------
# Hand-movement tracker
# ---------------------------------------------------------------------------
class HandMovementTracker:
    def __init__(self):
        self.enabled  = False
        self._lock    = threading.Lock()
        self._samples: list = []
        self._running = False
        if not NICLA_ENABLED:
            return
        try:
            from bleak import BleakClient, BleakScanner  # noqa
            self.enabled = True
        except ImportError:
            print("[HandMovementTracker] 'bleak' not installed – sensor disabled.")

    def start(self):
        if not self.enabled: return
        self._running = True
        threading.Thread(target=self._ble_loop, daemon=True).start()

    def stop(self):
        self._running = False

    def begin_question(self):
        with self._lock: self._samples.clear()

    def end_question(self) -> dict:
        with self._lock: s = list(self._samples)
        if not s:
            return {"sample_count": 0, "mean_magnitude": 0.0, "max_magnitude": 0.0}
        return {"sample_count": len(s),
                "mean_magnitude": round(sum(s) / len(s), 4),
                "max_magnitude":  round(max(s), 4)}

    def _ble_loop(self):
        import asyncio
        asyncio.run(self._async_ble_loop())

    async def _async_ble_loop(self):
        from bleak import BleakClient, BleakScanner
        while self._running:
            try:
                device = await BleakScanner.find_device_by_filter(
                    lambda d, _: NICLA_DEVICE_NAME.lower() in (d.name or "").lower(),
                    timeout=10.0)
                if device is None:
                    await __import__('asyncio').sleep(5); continue
                async with BleakClient(device) as client:
                    await client.start_notify(NICLA_CHAR_UUID, self._handler)
                    while self._running and client.is_connected:
                        await __import__('asyncio').sleep(0.1)
                    await client.stop_notify(NICLA_CHAR_UUID)
            except Exception:
                await __import__('asyncio').sleep(5)

    def _handler(self, _sender, data: bytearray):
        import struct
        try:
            if len(data) >= 6:
                ax, ay, az = struct.unpack_from('<hhh', data, 0)
                with self._lock:
                    self._samples.append(((ax/100)**2 + (ay/100)**2 + (az/100)**2) ** 0.5)
        except Exception: pass


# ---------------------------------------------------------------------------
# Hard Question Engine
# ---------------------------------------------------------------------------
class HardQuestionEngine:
    """
    All questions require pencil-and-paper working steps.
    Answers are always exact integers or simple fractions —
    no calculator needed, but no mental shortcuts either.
    """

    @staticmethod
    def _simultaneous() -> tuple:
        x = random.randint(2, 10)
        y = random.randint(2, 10)
        a = random.randint(2, 6)
        b = random.randint(1, 5)
        c = random.randint(1, 5)
        d = random.randint(2, 6)
        while a * d - b * c == 0:
            d = random.randint(2, 6)
        p = a * x + b * y
        q = c * x + d * y
        q_text = (f"Solve the system:\n"
                  f"  {a}x + {b}y = {p}\n"
                  f"  {c}x + {d}y = {q}\n"
                  f"What is the value of x?")
        return q_text, str(x)

    @staticmethod
    def _quadratic() -> tuple:
        r1 = random.randint(-8, -1)
        r2 = random.randint(1, 8)
        b = -(r1 + r2)
        c = r1 * r2
        b_str = f"+ {b}" if b >= 0 else f"- {abs(b)}"
        c_str = f"+ {c}" if c >= 0 else f"- {abs(c)}"
        larger_root = max(r1, r2)
        q_text = (f"Solve:  x² {b_str}x {c_str} = 0\n"
                  f"What is the larger root?")
        return q_text, str(larger_root)

    @staticmethod
    def _geometry() -> tuple:
        sub = random.choice(['composite', 'similar', 'pyth', 'rect_from_area'])

        if sub == 'composite':
            W  = random.randint(8, 16)
            H  = random.randint(8, 16)
            w  = random.randint(2, W - 2)
            h  = random.randint(2, H - 2)
            ans = W * H - w * h
            q_text = (f"An L-shaped figure is made by taking a {W}×{H} rectangle "
                      f"and removing a {w}×{h} rectangle from one corner.\n"
                      f"What is the area of the L-shape?")

        elif sub == 'similar':
            scale = random.randint(2, 4)
            a1    = random.randint(3, 10)
            b1    = random.randint(3, 10)
            a2    = a1 * scale
            b2    = b1 * scale
            ans   = b2
            q_text = (f"Two similar triangles have corresponding sides in proportion.\n"
                      f"Triangle 1 has sides {a1} and {b1}.\n"
                      f"Triangle 2 has a side of {a2} corresponding to the {a1} side.\n"
                      f"What is the side of Triangle 2 that corresponds to {b1}?")

        elif sub == 'pyth':
            triples = [(3,4,5),(5,12,13),(8,15,17),(7,24,25),(9,40,41),(6,8,10),(9,12,15)]
            a, b, c = random.choice(triples)
            mult    = random.randint(1, 3)
            a, b, c = a*mult, b*mult, c*mult
            ans     = c
            q_text  = (f"A right triangle has legs of length {a} and {b}.\n"
                       f"What is the length of the hypotenuse?")

        else:
            w = random.randint(4, 12)
            k = random.randint(2, 8)
            area = w * (w + k)
            perim = 2 * (w + w + k)
            ans   = perim
            q_text = (f"A rectangle has an area of {area} square units.\n"
                      f"Its length is {k} more than its width.\n"
                      f"What is its perimeter?")

        return q_text, str(ans)

    @staticmethod
    def _sequence_sum() -> tuple:
        a = random.randint(1, 10)
        d = random.randint(1, 8)
        n = random.randint(5, 12)
        S = n * (2 * a + (n - 1) * d) // 2
        q_text = (f"Find the sum of the first {n} terms of the arithmetic sequence "
                  f"that starts at {a} with a common difference of {d}.")
        return q_text, str(S)

    @staticmethod
    def _pct_chain() -> tuple:
        candidates = []
        for start in [100, 200, 400, 500, 1000]:
            for p1 in [10, 20, 25, 50]:
                for p2 in [10, 20, 25, 50]:
                    after1 = start * (100 + p1) // 100
                    after2 = after1 * (100 - p2) // 100
                    if after2 == start * (100 + p1) * (100 - p2) // 10000:
                        candidates.append((start, p1, p2, after2))
        start, p1, p2, ans = random.choice(candidates)
        q_text = (f"A price starts at ${start}. "
                  f"It increases by {p1}%, then decreases by {p2}%.\n"
                  f"What is the final price?")
        return q_text, str(ans)

    @staticmethod
    def _age_problem() -> tuple:
        sub = random.choice(['ratio_now', 'future_sum', 'diff_and_sum'])

        if sub == 'ratio_now':
            r1 = random.randint(1, 4)
            r2 = random.randint(r1 + 1, 7)
            k  = random.randint(3, 8)
            age_a = r1 * k
            age_b = r2 * k
            q_text = (f"Two people's ages are in the ratio {r1}:{r2}. "
                      f"The sum of their ages is {age_a + age_b}.\n"
                      f"How old is the older person?")
            ans = age_b

        elif sub == 'future_sum':
            a = random.randint(10, 30)
            b = random.randint(10, 30)
            t = random.randint(3, 10)
            s = (a + t) + (b + t)
            q_text = (f"Person A is currently {a} years old and Person B is {b}.\n"
                      f"In how many years will the sum of their ages be {s}?")
            ans = t

        else:
            diff = random.randint(2, 15)
            total = random.randint(20, 60)
            while (total + diff) % 2 != 0:
                total += 1
            older = (total + diff) // 2
            q_text = (f"The sum of two people's ages is {total}. "
                      f"One person is {diff} years older than the other.\n"
                      f"How old is the older person?")
            ans = older

        return q_text, str(ans)

    @staticmethod
    def _mixture() -> tuple:
        for _ in range(200):
            c1 = random.randint(60, 90)
            c2 = random.randint(10, 40)
            ct = random.randint(c2 + 5, c1 - 5)
            V  = random.randint(4, 20)
            num = V * (ct - c2)
            den = c1 - c2
            if num % den == 0:
                x = num // den
                q_text = (f"A chemist has a {c1}% solution and a {c2}% solution. "
                          f"How many litres of the {c1}% solution must be mixed "
                          f"with the {c2}% solution to produce {V} litres "
                          f"of a {ct}% solution?")
                return q_text, str(x)
        return HardQuestionEngine._sequence_sum()

    @staticmethod
    def _work_three() -> tuple:
        combos = [
            (2, 3, 6), (3, 4, 6), (2, 4, 4),
            (3, 6, 6), (4, 6, 12), (2, 6, 3),
            (6, 4, 12), (3, 3, 6),
        ]
        a, b, c = random.choice(combos)
        total_rate = fractions.Fraction(1, a) + fractions.Fraction(1, b) + fractions.Fraction(1, c)
        ans = str(fractions.Fraction(1, 1) / total_rate)
        q_text = (f"Worker A completes a job alone in {a} days, "
                  f"Worker B in {b} days, and Worker C in {c} days.\n"
                  f"How many days does it take all three working together?\n"
                  f"(Express as a fraction if needed)")
        return q_text, ans

    @staticmethod
    def _digit_problem() -> tuple:
        sub = random.choice(['reverse', 'sum_and_diff', 'divisibility'])

        if sub == 'reverse':
            t = random.randint(2, 7)
            u = random.randint(t + 1, 9)
            original = 10 * t + u
            reversed_num = 10 * u + t
            diff = reversed_num - original
            sum_digits = t + u
            q_text = (f"A two-digit number has a digit sum of {sum_digits}. "
                      f"When the digits are reversed, the new number is "
                      f"{diff} more than the original.\n"
                      f"What is the original number?")
            ans = original

        elif sub == 'sum_and_diff':
            t = random.randint(1, 7)
            u = random.randint(t + 1, 9)
            q_text = (f"The sum of the digits of a two-digit number is {t + u}. "
                      f"The difference between the digits is {u - t} "
                      f"(larger minus smaller).\n"
                      f"What is the larger of the two possible two-digit numbers?")
            ans = max(10 * t + u, 10 * u + t)

        else:
            base = random.randint(2, 8)
            mult = random.randint(3, 9)
            target = base * mult
            M     = random.randint(10, 60)
            start = (M // target + 1) * target
            q_text = (f"What is the smallest two-digit multiple of {target} "
                      f"that is greater than {M}?")
            ans = start

        return q_text, str(ans)

    @staticmethod
    def _proportion_word() -> tuple:
        sub = random.choice(['map_scale', 'recipe_scale', 'gear_ratio'])

        if sub == 'map_scale':
            scale_cm  = random.randint(2, 5)
            scale_km  = random.randint(10, 50)
            actual_cm = random.randint(3, 12)
            ans       = actual_cm * scale_km // scale_cm
            q_text    = (f"On a map, {scale_cm} cm represents {scale_km} km. "
                         f"Two cities are {actual_cm} cm apart on the map.\n"
                         f"How many km apart are they in real life?")

        elif sub == 'recipe_scale':
            orig_serves = random.choice([4, 6, 8])
            need_serves = orig_serves * random.randint(2, 4)
            ingredient  = random.randint(3, 12) * orig_serves // 2
            ans         = ingredient * need_serves // orig_serves
            q_text      = (f"A recipe for {orig_serves} servings uses "
                           f"{ingredient} grams of flour.\n"
                           f"How many grams of flour are needed for "
                           f"{need_serves} servings?")

        else:
            t1 = random.choice([12, 15, 18, 20, 24, 30])
            t2 = random.choice([6,  8,  9,  10, 12, 15])
            rpm1 = random.choice([60, 80, 100, 120, 150])
            rpm2 = t1 * rpm1 // t2
            q_text = (f"Gear A has {t1} teeth and spins at {rpm1} rpm. "
                      f"It meshes with Gear B which has {t2} teeth.\n"
                      f"How fast does Gear B spin (in rpm)?")
            ans = rpm2

        return q_text, str(ans)

    _GENERATORS = [
        _simultaneous.__func__,
        _quadratic.__func__,
        _geometry.__func__,
        _sequence_sum.__func__,
        _pct_chain.__func__,
        _age_problem.__func__,
        _mixture.__func__,
        _work_three.__func__,
        _digit_problem.__func__,
        _proportion_word.__func__,
    ]

    @classmethod
    def generate(cls) -> tuple:
        q_text, correct = random.choice(cls._GENERATORS)()
        return cls._pack(q_text, correct)

    @staticmethod
    def _pack(question: str, correct: str) -> tuple:
        choices = {correct}
        attempts = 0
        while len(choices) < 4 and attempts < 300:
            attempts += 1
            noise = HardQuestionEngine._nearby(correct)
            if noise not in choices:
                choices.add(noise)
        choices = list(choices)
        random.shuffle(choices)
        return question, correct, choices

    @staticmethod
    def _nearby(value_str: str) -> str:
        if '/' in value_str:
            try:
                f     = fractions.Fraction(value_str)
                delta = fractions.Fraction(random.randint(-3, 3),
                                           max(1, f.denominator))
                cand  = f + delta
                if cand > 0:
                    return str(cand)
            except Exception:
                pass
        try:
            val  = float(value_str)
            offs = [-8, -6, -5, -4, -3, -2, -1, 1, 2, 3, 4, 5, 6, 8,
                    int(val * 0.15) or 2, -(int(val * 0.15) or 2)]
            cand = round(val + random.choice(offs), 2)
            return str(int(cand)) if cand == int(cand) else str(cand)
        except Exception:
            return str(random.randint(2, 150))


# ---------------------------------------------------------------------------
# Test UI
# ---------------------------------------------------------------------------
class HardMathTest:
    def __init__(self, root, username, participant_data=None, callback=None):
        self.root     = root
        self.root.title(f"Math Test — {TEST_LABEL}")
        self.root.configure(bg="#0d0d0d")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        sw, sh = root.winfo_screenwidth(), root.winfo_screenheight()
        ww = max(900,  min(1920, int(sw * 0.90)))
        wh = max(650,  min(1080, int(sh * 0.90)))
        root.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")
        root.minsize(900, 650)
        root.resizable(True, True)
        self.scale = min(ww / 1920, wh / 1080)

        def fs(n): return max(8, int(n * self.scale))
        self._fs = fs

        # State
        self.running             = True
        self.username            = username
        self.participant_data    = participant_data or {}
        self.path_to_file        = (participant_data.get('_xlsx_path')
                                    or get_xlsx_path(username))
        self.correct_count       = 0
        self.wrong_count         = 0
        self.miss_count          = 0
        self.total_questions     = 0
        self.question_start_time = 0
        self.answered            = True
        self.current_timer_id    = None
        self.q_timer_id          = None
        self._hide_timer_id      = None
        self.callback            = callback
        self.question_text       = ""
        self.correct_answer      = ""
        self.choices: list       = []
        self.correct_option_idx  = 0
        self.music_playing       = False

        self.tracker = HandMovementTracker()
        self.tracker.start()

        self._build_ui(fs)

        self.start_time        = time.time()
        self.start_time_header = time.strftime('%m/%d/%Y %H:%M:%S',
                                               time.localtime(self.start_time))

        print("Saving data to:", os.path.abspath(self.path_to_file))
        append_timeline_event(self.path_to_file, "Test Start")

        self.countdown(COUNTDOWN)

    # ------------------------------------------------------------------ UI
    def _build_ui(self, fs):
        BG = "#0d0d0d"
        s  = self.scale

        top = tk.Frame(self.root, bg=BG)
        top.pack(fill=tk.X, padx=int(18*s), pady=int(12*s))

        self.q_count_label = tk.Label(top, text=f"Q 0 / {TOTAL_QUESTIONS}",
                                       bg=BG, fg="white",
                                       font=("Courier New", fs(20), "bold"),
                                       anchor='w')
        self.q_count_label.pack(side=tk.LEFT)

        self.score_label = tk.Label(top, text="✓ 0   ✗ 0",
                                    bg=BG, fg="#aaaaaa",
                                    font=("Courier New", fs(16)))
        self.score_label.pack(side=tk.RIGHT)

        bar_frame = tk.Frame(self.root, bg=BG)
        bar_frame.pack(fill=tk.X, padx=int(18*s), pady=(0, int(6*s)))

        self.q_time_label = tk.Label(bar_frame, text="", bg=BG, fg="#ffcc44",
                                     font=("Courier New", fs(14)))
        self.q_time_label.pack(side=tk.LEFT)

        self.bar_canvas = tk.Canvas(bar_frame, height=int(12*s),
                                    bg="#222222", highlightthickness=0)
        self.bar_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True,
                             padx=(int(10*s), 0))

        self.question_label = tk.Label(
            self.root, text="",
            font=("Georgia", fs(23)),
            bg=BG, fg="white",
            wraplength=int(0.80 * self.root.winfo_screenwidth()),
            justify="left",
            padx=int(20*s), pady=int(10*s))
        self.question_label.pack(pady=(int(16*s), int(6*s)), expand=True)

        self.hint_label = tk.Label(
            self.root,
            text="✏️  Write your steps on paper — select the closest answer.",
            bg=BG, fg="#555555",
            font=("Georgia", fs(13), "italic"))
        self.hint_label.pack()

        self.feedback_label = tk.Label(
            self.root, text="",
            font=("Courier New", fs(21), "bold"),
            bg=BG, fg="white")
        self.feedback_label.pack(pady=int(6*s))

        self.btn_bg      = "#1a1a2e"
        self.btn_fg      = "#e0e0e0"
        self.btn_hover   = "#16213e"
        self.btn_correct = "#2d6a4f"
        self.btn_wrong   = "#9b2226"

        self.buttons_frame = tk.Frame(self.root, bg=BG)
        self.option_buttons = []
        for i in range(4):
            btn = tk.Button(
                self.buttons_frame, text="",
                font=("Courier New", fs(21)),
                bg=self.btn_bg, fg=self.btn_fg,
                activebackground=self.btn_hover,
                activeforeground=self.btn_fg,
                relief=tk.FLAT, bd=0,
                padx=int(14*s), pady=int(10*s),
                width=22, wraplength=int(280*s),
                justify="center",
                command=lambda idx=i: self.check_answer(idx))
            self.option_buttons.append(btn)

    # ------------------------------------------------------------------ close
    def on_closing(self):
        self.running = False
        self.stop_music()
        self.tracker.stop()
        for tid in [self.current_timer_id, self.q_timer_id,
                    self._hide_timer_id]:
            if tid:
                try: self.root.after_cancel(tid)
                except Exception: pass
        time.sleep(0.1)
        try: self.root.destroy()
        except tk.TclError: pass
        if self.callback: self.callback()

    # ------------------------------------------------------------------ countdown
    def countdown(self, count):
        if not self.running: return
        if count > 0:
            try:
                self.question_label.config(
                    text=str(count),
                    font=("Courier New", self._fs(110), "bold"))
                self.root.after(1000, self.countdown, count - 1)
            except tk.TclError: pass
        else:
            try:
                self.question_label.config(
                    text="", font=("Georgia", self._fs(23)))
                self.buttons_frame.pack(pady=int(14 * self.scale))
                for i, btn in enumerate(self.option_buttons):
                    r, c = divmod(i, 2)
                    btn.grid(row=r, column=c,
                             padx=int(12*self.scale),
                             pady=int(8*self.scale))
                self._start_music_thread()
                self.generate_question()
            except tk.TclError: pass

    # ------------------------------------------------------------------ audio
    def _play_music(self):
        path = get_resource_path("clock.wav")
        self.music_playing = True
        if _SYSTEM == "Windows":
            try:
                import winsound
                winsound.PlaySound(None, winsound.SND_PURGE)
                while self.music_playing:
                    winsound.PlaySound(path, winsound.SND_FILENAME)
            except Exception as e:
                print(f"[Audio] {e}")
        else:
            while self.music_playing:
                if _SYSTEM == "Darwin":
                    cmds = [["afplay", path]]
                else:
                    cmds = [["aplay", "-q", path], ["paplay", path]]
                for cmd in cmds:
                    try:
                        proc = _subprocess.Popen(cmd,
                                                 stdout=_subprocess.DEVNULL,
                                                 stderr=_subprocess.DEVNULL)
                        while proc.poll() is None:
                            if not self.music_playing:
                                proc.terminate()
                                proc.wait()
                                return
                            time.sleep(0.1)
                        break
                    except FileNotFoundError:
                        continue
                else:
                    break

    def stop_music(self):
        self.music_playing = False
        if _SYSTEM == "Windows":
            try:
                import winsound
                winsound.PlaySound(None, winsound.SND_PURGE)
            except Exception: pass

    def _start_music_thread(self):
        threading.Thread(target=self._play_music, daemon=True).start()

    # ------------------------------------------------------------------ q bar
    def _start_q_bar(self, total_secs: int):
        self._q_bar_total   = total_secs
        self._q_bar_elapsed = 0
        self._update_q_bar()

    def _update_q_bar(self):
        if not self.running or self.answered: return
        try:
            elapsed  = self._q_bar_elapsed
            total    = self._q_bar_total
            fraction = max(0.0, 1.0 - elapsed / total)
            self.q_time_label.config(text=f"{total - elapsed}s")
            w = self.bar_canvas.winfo_width()
            h = self.bar_canvas.winfo_height()
            self.bar_canvas.delete("all")
            colour = ("#55b87a" if fraction > 0.4 else
                      "#ffcc44" if fraction > 0.2 else "#d9534f")
            self.bar_canvas.create_rectangle(
                0, 0, int(w * fraction), h, fill=colour, outline="")
            self._q_bar_elapsed += 1
            self.q_timer_id = self.root.after(1000, self._update_q_bar)
        except tk.TclError: pass

    # ------------------------------------------------------------------ questions
    def generate_question(self):
        if not self.running: return
        # End game if we've reached the question limit
        if self.total_questions >= TOTAL_QUESTIONS:
            self.end_game()
            return
        try:
            self.tracker.begin_question()
            self.question_start_time = time.time()

            q_text, correct, choices = HardQuestionEngine.generate()
            self.question_text      = q_text
            self.correct_answer     = correct
            self.choices            = choices
            self.correct_option_idx = choices.index(correct)

            self.question_label.config(text=q_text,
                                       font=("Georgia", self._fs(23)))
            self.feedback_label.config(text="")
            self._update_score_label()

            letters = ["A", "B", "C", "D"]
            for i, btn in enumerate(self.option_buttons):
                btn.config(text=f"{letters[i]}.  {choices[i]}",
                           bg=self.btn_bg, fg=self.btn_fg, state=tk.NORMAL)

            self.answered = False
            if self.q_timer_id:
                self.root.after_cancel(self.q_timer_id)
            if self._hide_timer_id:
                self.root.after_cancel(self._hide_timer_id)
            self._start_q_bar(QUESTION_TIME_SEC)
            self._hide_timer_id = self.root.after(
                QUESTION_TIME_SEC * 1000, self.hide_question)
        except tk.TclError: pass

    def _update_score_label(self):
        self.score_label.config(
            text=f"✓ {self.correct_count}   ✗ {self.wrong_count}")
        self.q_count_label.config(
            text=f"Q {self.total_questions} / {TOTAL_QUESTIONS}")

    # ------------------------------------------------------------------ answer
    def check_answer(self, idx):
        if not self.running or self.answered: return
        self.answered   = True
        response_ms     = int((time.time() - self.question_start_time) * 1000)
        movement        = self.tracker.end_question()
        for tid in [self._hide_timer_id, self.q_timer_id]:
            if tid:
                try: self.root.after_cancel(tid)
                except Exception: pass
        try:
            correct = (idx == self.correct_option_idx)
            self.total_questions += 1
            if correct:
                self.correct_count += 1
                self.option_buttons[idx].config(bg=self.btn_correct, fg="white")
                self.feedback_label.config(text="✓  Correct!", fg="#55b87a")
                self.write_result_to_xlsx("CORRECT", idx, response_ms, movement)
            else:
                self.wrong_count += 1
                self.option_buttons[idx].config(bg=self.btn_wrong, fg="white")
                self.option_buttons[self.correct_option_idx].config(
                    bg=self.btn_correct, fg="white")
                self.feedback_label.config(
                    text=f"✗  Incorrect — answer was {self.correct_answer}",
                    fg="#ff6b6b")
                self.write_result_to_xlsx("WRONG", idx, response_ms, movement)
            self._update_score_label()
            for btn in self.option_buttons:
                btn.config(state=tk.DISABLED)
            self.q_time_label.config(text="")
            self.root.after(1200, self._advance)
        except tk.TclError: pass

    def _advance(self):
        if not self.running: return
        try:
            for btn in self.option_buttons:
                btn.config(text="", bg=self.btn_bg, state=tk.NORMAL)
            self.feedback_label.config(text="")
            self.generate_question()
        except tk.TclError: pass

    def hide_question(self):
        if not self.running or self.answered: return
        self.answered   = True
        response_ms     = int((time.time() - self.question_start_time) * 1000)
        movement        = self.tracker.end_question()
        self.total_questions += 1
        self.miss_count      += 1
        self.write_result_to_xlsx("MISS", -1, response_ms, movement)
        try:
            self.feedback_label.config(
                text=f"⏱  Time's up! Answer: {self.correct_answer}",
                fg="#ffcc44")
            self.q_time_label.config(text="")
            self._update_score_label()
            for btn in self.option_buttons:
                btn.config(text="", bg=self.btn_bg)
            self.root.after(1200, self._advance)
        except tk.TclError: pass

    # ------------------------------------------------------------------ end
    def end_game(self):
        if not self.running: return
        for tid in [self._hide_timer_id, self.q_timer_id,
                    self.current_timer_id]:
            if tid:
                try: self.root.after_cancel(tid)
                except Exception: pass
        self.running  = False
        self.answered = True
        self.tracker.stop()
        try:
            self.question_label.config(
                text="Done!",
                font=("Courier New", self._fs(80), "bold"))
            self.feedback_label.config(text="")
            self.hint_label.config(text="")
            self.q_time_label.config(text="")
            for btn in self.option_buttons:
                btn.config(text="", state=tk.DISABLED)
        except tk.TclError: pass

        append_timeline_event(self.path_to_file, "Test End")
        self.write_summary_to_xlsx()
        self.stop_music()
        try: self.root.after(2500, self.root.destroy)
        except tk.TclError: pass
        if self.callback: self.callback()

    # ------------------------------------------------------------------ Excel
    def write_result_to_xlsx(self, result, idx, response_ms, movement):
        selected = self.option_buttons[idx].cget("text") if idx >= 0 else "MISS"
        row = [
            time.strftime('%m/%d/%Y %H:%M:%S', time.localtime()),
            int((time.time() - self.start_time) * 1000),
            response_ms, QUESTION_TIME_SEC,
            self.question_text.replace('\n', ' '),
            self.correct_answer, selected, result,
            movement["sample_count"], movement["mean_magnitude"],
            movement["max_magnitude"],
        ]
        threading.Thread(
            target=append_trial_to_xlsx,
            args=(self.path_to_file, row),
            daemon=True).start()

    def write_summary_to_xlsx(self):
        payout = max(0.0, MAX_PAYOUT - PENALTY_PER_WRONG * (self.wrong_count + self.miss_count))
        summary = {
            "Start Time":   self.start_time_header,
            "Total Q":      self.total_questions,
            "Correct":      self.correct_count,
            "Wrong":        self.wrong_count,
            "Miss":         self.miss_count,
            "Payout ($)":   round(payout, 2),
            "Test Version": TEST_LABEL,
        }
        write_summary_to_xlsx(self.path_to_file, summary)


# ---------------------------------------------------------------------------
# Ending Screen  (shown after post-test relaxation)
# ---------------------------------------------------------------------------
class EndingScreen:
    """
    Closing screen shown after the post-test relaxation period ends.
    Displays payout info and completion message.
    """

    BG        = "#0a0e1a"
    CARD_BG   = "#111828"
    BORDER    = "#1e3050"
    TEXT_HEAD = "#c8d8f0"
    TEXT_BODY = "#8aa8cc"
    TEXT_DIM  = "#4a6080"

    def __init__(self, master, participant_data=None):
        self.master = master
        self.master.title("Session Complete")
        self.master.configure(bg=self.BG)
        self.master.protocol("WM_DELETE_WINDOW", self.master.destroy)

        sw = master.winfo_screenwidth()
        sh = master.winfo_screenheight()
        self._base = min(sw, sh)
        ww = max(520, min(int(sw * 0.50), 780))
        wh = max(420, min(int(sh * 0.60), 600))
        x  = (sw - ww) // 2
        y  = (sh - wh) // 2
        master.geometry(f"{ww}x{wh}+{x}+{y}")
        master.resizable(True, True)
        self._ww = ww
        self._wh = wh

        # Read summary from xlsx to get payout if available
        self._participant_data = participant_data or {}
        self._build_ui()

    def _fs(self, pt):
        return max(8, int(pt * self._base / 1080))

    def _build_ui(self):
        outer = tk.Frame(self.master, bg=self.BG)
        outer.pack(fill=tk.BOTH, expand=True)

        card_wrap = tk.Frame(outer, bg=self.BORDER)
        card_wrap.place(relx=0.5, rely=0.5, anchor="center",
                        relwidth=0.88, relheight=0.90)

        card = tk.Frame(card_wrap, bg=self.CARD_BG)
        card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill=tk.BOTH, expand=True,
                   padx=int(self._ww * 0.08),
                   pady=int(self._wh * 0.07))

        # Checkmark icon
        tk.Label(
            inner, text="✓",
            font=("Helvetica Neue", self._fs(52)),
            bg=self.CARD_BG, fg="#4a90d9"
        ).pack(pady=(0, int(self._wh * 0.02)))

        tk.Label(
            inner,
            text="Session Complete",
            font=("Helvetica Neue", self._fs(28), "bold"),
            bg=self.CARD_BG, fg=self.TEXT_HEAD
        ).pack()

        tk.Frame(inner, bg=self.BORDER, height=1).pack(
            fill=tk.X, pady=(int(self._wh * 0.03), int(self._wh * 0.035)))

        # Try to load payout from xlsx
        payout_text = self._get_payout_text()

        tk.Label(
            inner,
            text=payout_text,
            font=("Helvetica Neue", self._fs(14)),
            bg=self.CARD_BG, fg=self.TEXT_BODY,
            justify="center",
            wraplength=int(self._ww * 0.72)
        ).pack()

        tk.Frame(inner, bg=self.CARD_BG).pack(fill=tk.BOTH, expand=True)

        tk.Label(
            inner,
            text="Thank you for participating. You may close this window.",
            font=("Helvetica Neue", self._fs(11), "italic"),
            bg=self.CARD_BG, fg=self.TEXT_DIM
        ).pack(pady=(0, int(self._wh * 0.01)))

    def _get_payout_text(self):
        xlsx = self._participant_data.get("_xlsx_path", "")
        correct = wrong = miss = 0
        if xlsx and os.path.exists(xlsx):
            try:
                from openpyxl import load_workbook
                wb = load_workbook(xlsx, read_only=True)
                ws = wb.active
                for row in ws.iter_rows(values_only=True):
                    if row and row[0] == "Correct":
                        correct = row[1] or 0
                    elif row and row[0] == "Wrong":
                        wrong = row[1] or 0
                    elif row and row[0] == "Miss":
                        miss = row[1] or 0
                wb.close()
            except Exception:
                pass

        payout = max(0.0, MAX_PAYOUT - PENALTY_PER_WRONG * (wrong + miss))
        total  = correct + wrong + miss

        lines = [
            f"You answered {correct} out of {total} questions correctly.",
            "",
            f"Your earnings for this session:  ${payout:.2f}",
            "",
            "Please let the researcher know you have finished.",
        ]
        return "\n".join(lines)


# ---------------------------------------------------------------------------
# Participant Intake UI
# ---------------------------------------------------------------------------
class ParticipantIntakeUI:

    GENDERS = ["— select —", "Male", "Female", "Non-binary", "Prefer not to say"]

    def __init__(self, root, on_submit=None):
        self.root      = root
        self.on_submit = on_submit

        self.root.title("Math Test — Participant Intake")
        self.root.configure(bg=IN_BG)
        self.root.resizable(True, True)

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        w  = 500
        h  = min(720, int(sh * 0.90))
        root.geometry("{}x{}+{}+{}".format(w, h, (sw - w) // 2, (sh - h) // 2))
        root.minsize(500, 600)

        self._vars = {}
        self._build_ui()

    def _build_ui(self):
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Intake.TCombobox",
                         fieldbackground=IN_ENTRY_BG,
                         background=IN_ENTRY_BG,
                         foreground=IN_TEXT,
                         bordercolor=IN_BORDER,
                         arrowcolor=IN_ACCENT,
                         padding=6)

        canvas = tk.Canvas(self.root, bg=IN_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll_frame  = tk.Frame(canvas, bg=IN_BG)
        scroll_window = canvas.create_window((0, 0), window=scroll_frame, anchor="nw")

        def _on_frame_configure(event):
            canvas.configure(scrollregion=canvas.bbox("all"))
        def _on_canvas_configure(event):
            canvas.itemconfig(scroll_window, width=event.width)
        scroll_frame.bind("<Configure>", _on_frame_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)

        outer = tk.Frame(scroll_frame, bg=IN_BG)
        outer.pack(fill=tk.BOTH, expand=True, padx=40, pady=36)

        tk.Label(outer, text="Participant Information",
                 bg=IN_BG, fg=IN_TEXT,
                 font=("Helvetica Neue", 20, "bold")).pack(anchor="w")
        tk.Label(outer,
                 text="Please fill in your details before starting the test.",
                 bg=IN_BG, fg=IN_MUTED,
                 font=("Helvetica Neue", 12)).pack(anchor="w", pady=(4, 20))

        tk.Frame(outer, bg=IN_BORDER, height=1).pack(fill=tk.X, pady=(0, 20))

        row = tk.Frame(outer, bg=IN_BG)
        row.pack(fill=tk.X, pady=(0, 12))
        self._field(row, "first_name", "First name *", side=tk.LEFT, width=18)
        tk.Frame(row, bg=IN_BG, width=16).pack(side=tk.LEFT)
        self._field(row, "last_name",  "Last name *",  side=tk.LEFT, width=18)

        self._field(outer, "email", "Email address *", full=True)
        self._field(outer, "age",   "Age *",           full=True, numeric=True)

        tk.Label(outer, text="Gender", bg=IN_BG, fg=IN_MUTED,
                 font=("Helvetica Neue", 12)).pack(anchor="w", pady=(0, 4))
        gender_var = tk.StringVar(value=self.GENDERS[0])
        self._vars["gender"] = gender_var
        ttk.Combobox(outer, textvariable=gender_var,
                     values=self.GENDERS, state="readonly",
                     style="Intake.TCombobox",
                     font=("Helvetica Neue", 13)).pack(fill=tk.X, pady=(0, 16))

        self._consent_var = tk.BooleanVar(value=False)
        consent_frame = tk.Frame(outer, bg="#f0f4f8",
                                  highlightthickness=1,
                                  highlightbackground=IN_BORDER)
        consent_frame.pack(fill=tk.X, pady=(0, 20))
        inner = tk.Frame(consent_frame, bg="#f0f4f8")
        inner.pack(padx=12, pady=10, fill=tk.X)
        tk.Checkbutton(inner, variable=self._consent_var,
                       bg="#f0f4f8", activebackground="#f0f4f8",
                       command=self._toggle_button).pack(side=tk.LEFT, anchor="n")
        tk.Label(inner,
                 text="I agree to participate in this study and consent to\n"
                      "the collection of my data for research purposes.",
                 bg="#f0f4f8", fg=IN_MUTED,
                 font=("Helvetica Neue", 11),
                 justify="left").pack(side=tk.LEFT, padx=(6, 0))

        self._error_label = tk.Label(outer, text="", bg=IN_BG, fg=IN_ERROR,
                                      font=("Helvetica Neue", 11))
        self._error_label.pack(anchor="w", pady=(0, 6))

        self._start_btn = tk.Button(
            outer, text="Start test →",
            bg="#aaaaaa", fg="#dddddd",
            font=("Helvetica Neue", 14, "bold"),
            activebackground="#16213e", activeforeground=IN_BTN_FG,
            relief=tk.FLAT, bd=0, padx=20, pady=12,
            state=tk.DISABLED,
            command=self._submit)
        self._start_btn.pack(fill=tk.X)

    def _field(self, parent, key, label, side=None, width=None,
               full=False, numeric=False):
        container = tk.Frame(parent, bg=IN_BG)
        if full:
            container.pack(fill=tk.X, pady=(0, 12))
        else:
            container.pack(side=side, fill=tk.X, expand=True)

        tk.Label(container, text=label, bg=IN_BG, fg=IN_MUTED,
                 font=("Helvetica Neue", 12)).pack(anchor="w", pady=(0, 4))

        var = tk.StringVar()
        self._vars[key] = var

        vcmd = None
        if numeric:
            vcmd = (container.register(
                lambda P: P == "" or (P.isdigit() and len(P) <= 3)), "%P")

        kw = {}
        if width:
            kw["width"] = width
        if vcmd:
            kw["validate"] = "key"
            kw["validatecommand"] = vcmd

        entry = tk.Entry(
            container, textvariable=var,
            bg=IN_ENTRY_BG, fg=IN_TEXT,
            font=("Helvetica Neue", 13),
            relief=tk.FLAT, bd=0,
            highlightthickness=1,
            highlightbackground=IN_BORDER,
            highlightcolor=IN_ACCENT,
            insertbackground=IN_TEXT,
            **kw)
        entry.pack(fill=tk.X if full else tk.NONE, expand=full, ipady=7)

    def _toggle_button(self):
        if self._consent_var.get():
            self._start_btn.config(state=tk.NORMAL, bg=IN_ACCENT, fg=IN_BTN_FG)
        else:
            self._start_btn.config(state=tk.DISABLED, bg="#aaaaaa", fg="#dddddd")

    def _validate(self):
        v = self._vars
        first = v["first_name"].get().strip()
        last  = v["last_name"].get().strip()
        email = v["email"].get().strip()
        age   = v["age"].get().strip()

        if not first or not last:
            self._error_label.config(text="Please enter your first and last name.")
            return False
        if not re.match(r"^[^\s@]+@[^\s@]+\.[^\s@]+$", email):
            self._error_label.config(text="Please enter a valid email address.")
            return False
        if not age or not (10 <= int(age) <= 99):
            self._error_label.config(text="Please enter a valid age (10–99).")
            return False
        self._error_label.config(text="")
        return True

    def _submit(self):
        if not self._validate():
            return
        v = self._vars
        gender = v["gender"].get()
        data = {
            "first_name": v["first_name"].get().strip(),
            "last_name":  v["last_name"].get().strip(),
            "username":   (v["first_name"].get().strip() + "_" +
                           v["last_name"].get().strip()).lower().replace(" ", "_"),
            "email":      v["email"].get().strip(),
            "age":        int(v["age"].get().strip()),
            "gender":     gender if gender != "— select —" else "",
        }
        self.root.destroy()
        if self.on_submit:
            self.on_submit(data)


# ---------------------------------------------------------------------------
# Session orchestrator
# ---------------------------------------------------------------------------

def run_pre_relaxation(participant_data):
    """Step 2: pre-test 3-minute relaxation.
    Excel file is created here so relaxation timestamps can be recorded.
    """
    xlsx = get_xlsx_path(participant_data["username"])
    create_xlsx_skeleton(xlsx, participant_data)
    participant_data["_xlsx_path"] = xlsx

    root = tk.Tk()
    RelaxationScreen(
        root,
        duration=RELAX_DURATION,
        on_done=lambda: show_instructions(participant_data),
        xlsx_path=xlsx,
        period_label="Pre-Test Relaxation",
        play_open_eyes=True,
    )
    root.mainloop()


def show_instructions(participant_data):
    """Step 3: instructions card with Start button."""
    root = tk.Tk()
    InstructionsWindow(
        root,
        on_ready=lambda: launch_test(participant_data),
    )
    root.mainloop()


def launch_test(participant_data):
    """Step 4: hard math test (10 questions)."""
    root = tk.Tk()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    ww = max(900,  min(1920, int(sw * 0.90)))
    wh = max(650,  min(1080, int(sh * 0.90)))
    root.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")
    root.minsize(900, 650)
    HardMathTest(
        root,
        username=participant_data["username"],
        participant_data=participant_data,
        callback=lambda: run_post_relaxation(participant_data),
    )
    root.mainloop()


def run_post_relaxation(participant_data):
    """Step 5: post-test 3-minute relaxation (no OpenEyes audio at end)."""
    xlsx = participant_data.get("_xlsx_path")
    root = tk.Tk()
    RelaxationScreen(
        root,
        duration=RELAX_DURATION,
        on_done=lambda: show_ending_screen(participant_data),
        xlsx_path=xlsx,
        period_label="Post-Test Relaxation",
        play_open_eyes=False,
    )
    root.mainloop()


def show_ending_screen(participant_data):
    """Step 6: ending screen with payout summary."""
    root = tk.Tk()
    EndingScreen(root, participant_data=participant_data)
    root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    intake_root = tk.Tk()
    ParticipantIntakeUI(intake_root, on_submit=run_pre_relaxation)
    intake_root.mainloop()