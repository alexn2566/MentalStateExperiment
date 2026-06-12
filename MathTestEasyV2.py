"""
Math Test - Version A: Easy / Mental Math
==========================================
Tier 1: Pure instant recall (12 seconds per question)
Tier 2: Quick mental math, 3-5 seconds of thought (18 seconds per question)

No scratch work needed. Adapts between tiers based on performance.
Participant intake screen shown before the test begins.

Flow:
  1. Intake screen (participant info)
  2. Pre-test relaxation screen (3 min, relax.mp3)
  3. "Open eyes" sound plays, then instructions window
  4. Math test (3 min)
  5. Post-test relaxation screen (3 min, relax.mp3)
"""

import tkinter as tk
from tkinter import ttk
import random
import csv
import time
import threading
import os
import sys
import re

# ---------------------------------------------------------------------------
# Utility
# ---------------------------------------------------------------------------
def get_resource_path(relative_path):
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------
TOTAL_GAME_TIME   = 60 * 3
COUNTDOWN         = 5
BASE_DATA_PATH    = './data/'
QUESTION_TIME     = {1: 10, 2: 15}
PROMOTE_THRESHOLD = 2
DEMOTE_THRESHOLD  = 2
NICLA_ENABLED     = False
NICLA_CHAR_UUID   = "19b10001-e8f2-537e-4f6c-d104768a1214"
NICLA_DEVICE_NAME = "NiclaSenseME"
TEST_LABEL        = "Easy / Mental Math"
XLSX_PREFIX       = "mathTest_easy_"

RELAX_DURATION    = 180   # 3 minutes

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

# Global state for the loop thread so we can stop it cleanly
_loop_stop_evt = None   # threading.Event when active
_loop_thread   = None   # threading.Thread when active


# _active_audio_proc holds a reference to a running subprocess (macOS/Linux)
# so _stop_all_audio() can kill it immediately from any thread.
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
    The relax screen sets stop_event after exactly RELAX_DURATION seconds
    via threading.Timer, guaranteeing a hard cut regardless of Tk timing.
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
                    break   # played successfully; loop again
                except FileNotFoundError:
                    continue
            else:
                print("[Audio] no suitable player found for loop")
                return


def _play_audio_once(path):
    """Play a .wav file exactly once (blocking). Run in a daemon thread."""
    _stop_all_audio()   # ensure no other audio is playing first
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
            "Tier", "Q Time Allowed (s)", "Question",
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

        # Set reasonable column widths
        col_widths = [20, 20, 20, 8, 18, 50, 16, 18, 12, 12, 14, 14]
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
    """
    Write a summary block below the trial data:
    Total Q / Correct / Wrong / Miss / Final Tier.
    """
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
    Calming full-screen relaxation window.

    Design:
      • Deep navy background (#0a0e1a) that fills the whole screen
      • Animated breathing circle: slow expand/contract cycle with a
        soft glow, guiding the participant's breath
      • "Breathe in … Breathe out …" cue text synced to the animation
      • Minimal, softly-lit countdown in the corner so it doesn't distract
      • Looping relax.mp3 in background
      • Works on macOS, Windows, and Linux (no platform-specific calls)
    """

    # Palette
    BG          = "#0a0e1a"
    RING_OUTER  = "#1a2a4a"
    RING_INNER  = "#162038"
    GLOW_COLOR  = "#4a90d9"
    TEXT_MAIN   = "#c8d8f0"
    TEXT_DIM    = "#4a6080"
    TEXT_BREATH = "#7ab0e0"

    # Breathing animation timing (ms)
    INHALE_MS   = 4000
    HOLD_MS     = 1500
    EXHALE_MS   = 5000
    REST_MS     = 1000

    def __init__(self, master, duration=RELAX_DURATION, on_done=None,
                 xlsx_path=None, period_label="Relaxation"):
        self.master        = master
        self.duration      = duration
        self.on_done       = on_done
        self.xlsx_path     = xlsx_path     # if set, timestamps are written here
        self.period_label  = period_label  # e.g. "Pre-Test Relaxation"
        self.running       = True
        self._stop_evt     = threading.Event()

        self.master.title("Relaxation")
        self.master.protocol("WM_DELETE_WINDOW", self.on_close)
        self.master.configure(bg=self.BG)

        # ---- Cross-platform fullscreen ----
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

        # Breathing animation state (circle only — no on-screen cue text)
        self._breath_phase  = "inhale"
        self._breath_frac   = 0.0
        self._anim_step_ms  = 40
        self._glow_items    = []

        # Responsive sizing helpers
        self._base = min(sw, sh)
        self._cx   = sw // 2
        self._cy   = int(sh * 0.46)

        # Record relaxation start timestamp
        if self.xlsx_path and os.path.exists(self.xlsx_path):
            threading.Thread(
                target=append_timeline_event,
                args=(self.xlsx_path, f"{self.period_label} Start"),
                daemon=True).start()

        # Start looping relax music
        music_path = get_resource_path('relax.wav')
        threading.Thread(
            target=_play_audio_loop,
            args=(music_path, self._stop_evt),
            daemon=True).start()

        # Hard audio timer: stop relax music and play OpenEyes.wav after
        # exactly RELAX_DURATION seconds — independent of Tk's scheduler.
        self._audio_timer = threading.Timer(
            float(duration), self._on_audio_timeout)
        self._audio_timer.daemon = True
        self._audio_timer.start()

        self._build_ui()

        # Hidden admin shortcut: right-arrow skips the relaxation immediately.
        self.master.bind("<Right>", lambda e: self._skip())

        self._tick_second()
        self._breath_step()

    # ------------------------------------------------------------------
    def _fs(self, base_pt):
        """Scale a font size relative to the shorter screen dimension."""
        return max(8, int(base_pt * self._base / 1080))

    def _build_ui(self):
        b  = self._base
        sw = self.sw
        sh = self.sh

        # Full-window canvas for drawing
        self.canvas = tk.Canvas(
            self.master, bg=self.BG,
            highlightthickness=0, bd=0)
        self.canvas.place(x=0, y=0, relwidth=1, relheight=1)

        # ---- Background concentric rings (decorative, static) ----
        cx, cy = self._cx, self._cy
        for r_frac, alpha_col in [
            (0.42, "#111828"), (0.35, "#121a30"),
            (0.28, "#131c34"), (0.21, "#141e38"),
        ]:
            r = int(b * r_frac)
            self.canvas.create_oval(
                cx - r, cy - r, cx + r, cy + r,
                fill=alpha_col, outline="")

        # ---- Breathing circle (will be redrawn each frame) ----
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

        # ---- Countdown (bottom-right corner, subtle) ----
        pad = int(b * 0.03)
        self._countdown_text = self.canvas.create_text(
            sw - pad, sh - pad,
            text=self._fmt_time(self.duration),
            font=(self._font("mono"), self._fs(15)),
            fill=self.TEXT_DIM,
            anchor="se")

        # ---- Soft footer note ----
        self.canvas.create_text(
            cx, int(sh * 0.93),
            text="You will be notified when it is time to continue.",
            font=(self._font("light"), self._fs(13)),
            fill=self.TEXT_DIM,
            anchor="center")

    # ------------------------------------------------------------------
    @staticmethod
    def _font(style="regular"):
        """Return a cross-platform font name for the requested weight."""
        fonts = {
            "thin":    ("Helvetica Neue", "Segoe UI Light", "Ubuntu Light",  "Helvetica"),
            "light":   ("Helvetica Neue", "Segoe UI",       "Ubuntu",        "Helvetica"),
            "regular": ("Helvetica Neue", "Segoe UI",       "Ubuntu",        "Arial"),
            "mono":    ("Courier New",    "Consolas",        "DejaVu Sans Mono", "Courier"),
        }
        # tkinter will fall back gracefully; just return the first candidate
        return fonts.get(style, ("Arial",))[0]

    @staticmethod
    def _fmt_time(secs):
        m, s = divmod(secs, 60)
        return f"{m}:{s:02d}"

    # ------------------------------------------------------------------
    def _lerp_color(self, frac):
        """Interpolate glow colour from dim blue → bright teal."""
        r0, g0, b0 = 0x1a, 0x3a, 0x6a   # dim
        r1, g1, b1 = 0x50, 0xb8, 0xe8   # bright
        r = int(r0 + (r1 - r0) * frac)
        g = int(g0 + (g1 - g0) * frac)
        b = int(b0 + (b1 - b0) * frac)
        return f"#{r:02x}{g:02x}{b:02x}"

    def _draw_breath_circle(self, frac):
        """Redraw the animated breathing circle at expansion fraction frac."""
        cx, cy = self._cx, self._cy
        r      = self._min_r + int((self._max_r - self._min_r) * frac)
        color  = self._lerp_color(frac)

        # Remove old glow layers
        for item in self._glow_items:
            self.canvas.delete(item)
        self._glow_items = []
        if self._circle_item:
            self.canvas.delete(self._circle_item)

        # Soft glow rings (outermost → innermost)
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

        # Main circle
        self._circle_item = self.canvas.create_oval(
            cx - r, cy - r, cx + r, cy + r,
            fill=color, outline="")

        # Small highlight dot
        hr = max(3, int(r * 0.18))
        hx = cx - int(r * 0.38)
        hy = cy - int(r * 0.38)
        self._glow_items.append(
            self.canvas.create_oval(
                hx - hr, hy - hr, hx + hr, hy + hr,
                fill="#a8d8f8", outline=""))

    # ------------------------------------------------------------------
    def _breath_step(self):
        """Animate the breathing circle using a sine-eased expand/contract."""
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
        else:  # rest
            self._breath_phase = "inhale"
            try:
                self._draw_breath_circle(self._breath_frac)
            except tk.TclError:
                return
            self.master.after(self.REST_MS, self._breath_step)
            return

        import math
        eased = (1 - math.cos(self._breath_frac * math.pi)) / 2

        try:
            self._draw_breath_circle(eased)
        except tk.TclError:
            return

        self.master.after(self._anim_step_ms, self._breath_step)

    # ------------------------------------------------------------------
    def _tick_second(self):
        if not self.running:
            return
        if self.duration > 0:
            try:
                self.canvas.itemconfig(
                    self._countdown_text,
                    text=self._fmt_time(self.duration))
                self.duration -= 1
                # Guard: only schedule next tick if the window still exists
                try:
                    self.master.after(1000, self._tick_second)
                except tk.TclError:
                    pass
            except tk.TclError:
                pass
        else:
            self._finish()

    # ------------------------------------------------------------------
    def _on_audio_timeout(self):
        """
        Called by threading.Timer after exactly RELAX_DURATION seconds.
        Stops the relax loop and plays OpenEyes.wav on a background thread.
        Does NOT touch Tk — safe to call from any thread.
        """
        self._stop_evt.set()   # stop relax loop
        # Play wake-up sound on a separate thread so we don't block
        eyes_path = get_resource_path('OpenEyes.wav')
        threading.Thread(target=_play_audio_once, args=(eyes_path,),
                         daemon=True).start()

    def _skip(self):
        """Right-arrow admin shortcut: cancel the audio timer and finish immediately."""
        if hasattr(self, '_audio_timer'):
            self._audio_timer.cancel()
        self._stop_evt.set()
        self._finish()

    # ------------------------------------------------------------------
    def _finish(self):
        if not self.running:
            return
        self.running = False
        self._stop_evt.set()
        # Cancel the audio timer if it hasn't fired yet (e.g. skip key pressed)
        if hasattr(self, '_audio_timer'):
            self._audio_timer.cancel()
        # Schedule the callback BEFORE destroying so the mainloop is still alive.
        if self.on_done:
            try:
                self.master.after(50, self._do_finish)
            except tk.TclError:
                self._do_finish()
        else:
            self._do_finish()

    def _do_finish(self):
        cb = self.on_done
        self.on_done = None   # prevent double-fire
        # Record relaxation end timestamp
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
        self.on_done = None   # user closed manually; skip the callback
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
    Plays the 'open eyes' sound, then shows exam instructions on a
    calming dark canvas that transitions naturally from the relaxation screen.
    Fully responsive — adapts to any screen size on Mac / Windows / Linux.
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

        # Responsive card size: 52% wide, 68% tall, centered
        ww = max(520, min(int(sw * 0.52), 820))
        wh = max(480, min(int(sh * 0.68), 700))
        x  = (sw - ww) // 2
        y  = (sh - wh) // 2
        self.master.geometry(f"{ww}x{wh}+{x}+{y}")
        self.master.resizable(True, True)

        self._ww = ww
        self._wh = wh

        # Play "open eyes" sound
        eyes_path = get_resource_path('OpenEyes.wav')
        threading.Thread(
            target=_play_audio_once,
            args=(eyes_path,),
            daemon=True).start()

        self._build_ui()

    def _fs(self, pt):
        return max(8, int(pt * self._base / 1080))

    def _build_ui(self):
        fs = self._fs

        # Outer dark background frame
        outer = tk.Frame(self.master, bg=self.BG)
        outer.pack(fill=tk.BOTH, expand=True)

        # Centered card frame with border effect
        card_wrap = tk.Frame(outer, bg=self.BORDER)
        card_wrap.place(relx=0.5, rely=0.5, anchor="center",
                        relwidth=0.88, relheight=0.90)

        card = tk.Frame(card_wrap, bg=self.CARD_BG)
        card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill=tk.BOTH, expand=True,
                   padx=int(self._ww * 0.07),
                   pady=int(self._wh * 0.06))

        # Header
        tk.Label(
            inner,
            text="Time to Wake Up",
            font=("Helvetica Neue", fs(26), "bold"),
            bg=self.CARD_BG, fg=self.TEXT_HEAD
        ).pack(anchor="w")

        tk.Label(
            inner,
            text="Your math test is about to begin.",
            font=("Helvetica Neue", fs(13)),
            bg=self.CARD_BG, fg=self.TEXT_DIM
        ).pack(anchor="w", pady=(2, 0))

        # Divider
        tk.Frame(inner, bg=self.BORDER, height=1).pack(
            fill=tk.X, pady=(int(self._wh * 0.025), int(self._wh * 0.03)))

        # Instruction bullets
        bullets = [
            ("Stay vigilant",
             "This is a practice exam — please give it your full effort."),
            ("Four choices per question",
             "Each question shows options A, B, C, and D. Tap the correct one."),
            ("Answer quickly",
             "Questions advance automatically when the timer runs out."),
            ("3 minutes total",
             "Work at a steady pace — accuracy and speed both matter."),
        ]

        for title, body in bullets:
            row = tk.Frame(inner, bg=self.CARD_BG)
            row.pack(fill=tk.X, pady=(0, int(self._wh * 0.018)))

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

        # Spacer
        tk.Frame(inner, bg=self.CARD_BG).pack(fill=tk.BOTH, expand=True)

        # Begin button
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
        self._samples = []
        self._running = False
        if not NICLA_ENABLED:
            return
        try:
            from bleak import BleakClient, BleakScanner  # noqa
            self.enabled = True
        except ImportError:
            print("[HandMovementTracker] 'bleak' not installed - sensor disabled.")

    def start(self):
        if not self.enabled:
            return
        self._running = True
        threading.Thread(target=self._ble_loop, daemon=True).start()

    def stop(self):
        self._running = False

    def begin_question(self):
        with self._lock:
            self._samples.clear()

    def end_question(self):
        with self._lock:
            s = list(self._samples)
        if not s:
            return {"sample_count": 0, "mean_magnitude": 0.0, "max_magnitude": 0.0}
        return {
            "sample_count":   len(s),
            "mean_magnitude": round(sum(s) / len(s), 4),
            "max_magnitude":  round(max(s), 4),
        }

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
                    await __import__('asyncio').sleep(5)
                    continue
                async with BleakClient(device) as client:
                    await client.start_notify(NICLA_CHAR_UUID, self._handler)
                    while self._running and client.is_connected:
                        await __import__('asyncio').sleep(0.1)
                    await client.stop_notify(NICLA_CHAR_UUID)
            except Exception:
                await __import__('asyncio').sleep(5)

    def _handler(self, _sender, data):
        import struct
        try:
            if len(data) >= 6:
                ax, ay, az = struct.unpack_from('<hhh', data, 0)
                with self._lock:
                    self._samples.append(
                        ((ax / 100) ** 2 + (ay / 100) ** 2 + (az / 100) ** 2) ** 0.5)
        except Exception:
            pass


# ---------------------------------------------------------------------------
# Question Engine
# ---------------------------------------------------------------------------
class QuestionEngine:

    @staticmethod
    def _make_choices(correct):
        choices = {correct}
        try:
            val  = float(correct)
            offs = [-6, -5, -4, -3, -2, -1, 1, 2, 3, 4, 5, 6,
                    -10, 10, -8, 8, -15, 15, -20, 20]
            random.shuffle(offs)
            for off in offs:
                if len(choices) >= 4:
                    break
                cand = int(val + off)
                if cand > 0:
                    choices.add(str(cand))
        except ValueError:
            pass
        safety = 0
        while len(choices) < 4 and safety < 500:
            safety += 1
            cand = str(random.randint(1, 200))
            choices.add(cand)
        result = list(choices)[:4]
        random.shuffle(result)
        return result

    @staticmethod
    def _tier1():
        kind = random.choice(['add_sub', 'add_sub', 'mult', 'double_half',
                              'pattern', 'compare'])
        if kind == 'add_sub':
            a  = random.randint(5, 50)
            b  = random.randint(5, 50)
            op = random.choice(['+', '-'])
            if op == '-' and b > a:
                a, b = b, a
            ans = a + b if op == '+' else a - b
            q   = "What is {} {} {}?".format(a, op, b)
        elif kind == 'mult':
            a   = random.randint(2, 12)
            b   = random.randint(2, 12)
            ans = a * b
            q   = "What is {} x {}?".format(a, b)
        elif kind == 'double_half':
            if random.random() < 0.5:
                n   = random.randint(5, 50)
                ans = n * 2
                q   = "What is double {}?".format(n)
            else:
                n   = random.choice([10, 12, 14, 16, 18, 20, 24,
                                     30, 40, 50, 60, 80, 100])
                ans = n // 2
                q   = "What is half of {}?".format(n)
        elif kind == 'pattern':
            start = random.randint(1, 10)
            step  = random.randint(2, 7)
            seq   = [start + i * step for i in range(4)]
            ans   = seq[-1] + step
            q     = ("What is the next number?\n"
                     "{},  {},  {},  {},  ___".format(*seq))
        else:
            a = random.randint(10, 99)
            b = random.randint(10, 99)
            while b == a:
                b = random.randint(10, 99)
            ans = max(a, b)
            q   = "Which is larger:  {}  or  {}?".format(a, b)
        choices = QuestionEngine._make_choices(str(ans))
        return q, str(ans), choices

    @staticmethod
    def _tier2():
        kind = random.choice(['two_step', 'pct', 'mult_carry',
                              'complement', 'word'])
        if kind == 'two_step':
            a   = random.randint(3, 15)
            b   = random.randint(3, 15)
            c   = random.randint(2, 10)
            op1 = random.choice(['+', '-', '*'])
            op2 = random.choice(['+', '-'])
            mid = int(eval("{} {} {}".format(a, op1, b)))
            ans = int(eval("{} {} {}".format(mid, op2, c)))
            q   = "What is {} {} {} {} {}?".format(a, op1, b, op2, c)
        elif kind == 'pct':
            whole = random.choice([10, 20, 30, 40, 50, 60, 80, 100, 200])
            pct   = random.choice([10, 20, 25, 50])
            ans   = int(whole * pct / 100)
            q     = "What is {}% of {}?".format(pct, whole)
        elif kind == 'mult_carry':
            a   = random.randint(2, 9)
            b   = random.randint(11, 19)
            ans = a * b
            q   = "What is {} x {}?".format(a, b)
        elif kind == 'complement':
            if random.random() < 0.5:
                n   = random.randint(11, 89)
                ans = 100 - n
                q   = "What is 100 - {}?".format(n)
            else:
                n   = random.choice([100, 200, 250, 300, 400,
                                     500, 600, 750, 800, 900])
                ans = 1000 - n
                q   = "What is 1000 - {}?".format(n)
        else:
            sub = random.choice(['items', 'age'])
            if sub == 'items':
                price = random.choice([5, 10, 15, 20, 25])
                count = random.randint(2, 8)
                ans   = price * count
                q     = ("Each item costs ${}.\n"
                         "How much do {} items cost?").format(price, count)
            else:
                age_now = random.randint(8, 25)
                years   = random.randint(2, 10)
                ans     = age_now + years
                q       = ("Someone is {} years old now.\n"
                           "How old will they be in {} years?").format(age_now, years)
        choices = QuestionEngine._make_choices(str(ans))
        return q, str(ans), choices

    @classmethod
    def generate(cls, tier, seen=None, max_attempts=200):
        """
        Generate a question for the given tier that hasn't been seen before.
        `seen` is a set of question strings maintained by the caller.
        If every possible question has been seen (extremely unlikely in 3 min),
        the seen set is cleared and generation continues fresh.
        """
        if seen is None:
            seen = set()
        for _ in range(max_attempts):
            if tier == 1:
                q, ans, choices = cls._tier1()
            else:
                q, ans, choices = cls._tier2()
            key = q.replace('\n', ' ')
            if key not in seen:
                seen.add(key)
                return q, ans, choices
        # Safety valve: clear history and return whatever comes next
        seen.clear()
        if tier == 1:
            return cls._tier1()
        return cls._tier2()


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

        var  = tk.StringVar()
        self._vars[key] = var

        vcmd = None
        if numeric:
            vcmd = (container.register(
                lambda P: P == "" or (P.isdigit() and len(P) <= 3)), '%P')

        entry = tk.Entry(
            container, textvariable=var,
            bg=IN_ENTRY_BG, fg=IN_TEXT,
            font=("Helvetica Neue", 13),
            relief=tk.FLAT, bd=0,
            highlightthickness=1,
            highlightbackground=IN_BORDER,
            highlightcolor=IN_ACCENT,
            insertbackground=IN_TEXT,
            **({"width": width} if width else {}),
            **({"validate": "key", "validatecommand": vcmd} if vcmd else {}))
        entry.pack(fill=tk.X if full else tk.NONE, expand=full, ipady=7)

    def _toggle_button(self):
        if self._consent_var.get():
            self._start_btn.config(state=tk.NORMAL, bg=IN_ACCENT, fg=IN_BTN_FG)
        else:
            self._start_btn.config(state=tk.DISABLED, bg="#aaaaaa", fg="#dddddd")

    def _validate(self):
        v     = self._vars
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
        v      = self._vars
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
# Math Test UI
# ---------------------------------------------------------------------------
class MathTestUI:

    TIER_INFO = {
        1: ("Tier 1 - Easy",   "#1e3a5f", "#7ec8e3"),
        2: ("Tier 2 - Medium", "#1e4d2b", "#78c76e"),
    }

    def __init__(self, root, username, participant_data=None, callback=None):
        self.root             = root
        self.root.title("Math Test - {}".format(TEST_LABEL))
        self.root.configure(bg="#0d0d0d")
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

        sw = root.winfo_screenwidth()
        sh = root.winfo_screenheight()
        ww = max(900,  min(1920, int(sw * 0.90)))
        wh = max(650,  min(1080, int(sh * 0.90)))
        root.geometry("{}x{}+{}+{}".format(ww, wh, (sw - ww) // 2, (sh - wh) // 2))
        root.minsize(900, 650)
        root.resizable(True, True)
        self.scale = min(ww / 1920, wh / 1080)

        self._fs = lambda n: max(8, int(n * self.scale))

        # Game state
        self.running             = True
        self.username            = username
        self.participant_data    = participant_data or {}
        self.path_to_file        = (participant_data.get('_xlsx_path')
                                    or get_xlsx_path(username))
        self.correct_count       = 0
        self.wrong_count         = 0
        self.miss_count          = 0
        self.total_questions     = 0
        self.remaining_time      = TOTAL_GAME_TIME
        self.question_start_time = 0
        self.answered            = True
        self.current_timer_id    = None
        self.q_timer_id          = None
        self._hide_timer_id      = None
        self.callback            = callback
        self.tier                = 1
        self.streak_correct      = 0
        self.streak_wrong        = 0
        self.current_q_time      = QUESTION_TIME[1]
        self.question_text       = ""
        self.correct_answer      = ""
        self.choices             = []
        self.correct_option_idx  = 0
        self.music_playing       = False
        self.seen_questions      = set()   # tracks question text to prevent repeats

        self.tracker = HandMovementTracker()
        self.tracker.start()

        self._build_ui()

        self.start_time        = time.time()
        self.start_time_header = time.strftime('%m/%d/%Y %H:%M:%S',
                                               time.localtime(self.start_time))

        # Excel file was already created by run_pre_relaxation; just log test start.
        print("Saving data to:", os.path.abspath(self.path_to_file))
        append_timeline_event(self.path_to_file, "Test Start")

        self.countdown(COUNTDOWN)

    # ------------------------------------------------------------------
    def _build_ui(self):
        bg = "#0d0d0d"
        s  = self.scale
        fs = self._fs

        top = tk.Frame(self.root, bg=bg)
        top.pack(fill=tk.X, padx=int(18 * s), pady=int(12 * s))

        self.time_label = tk.Label(
            top, text="Time: {}".format(self.remaining_time),
            bg=bg, fg="white",
            font=("Courier New", fs(20), "bold"), anchor='w')
        self.time_label.pack(side=tk.LEFT)

        self.score_label = tk.Label(
            top, text="correct: 0   wrong: 0",
            bg=bg, fg="#aaaaaa",
            font=("Courier New", fs(16)))
        self.score_label.pack(side=tk.RIGHT)

        bar_frame = tk.Frame(self.root, bg=bg)
        bar_frame.pack(fill=tk.X, padx=int(18 * s), pady=(0, int(6 * s)))

        self.q_time_label = tk.Label(
            bar_frame, text="", bg=bg, fg="#ffcc44",
            font=("Courier New", fs(14)))
        self.q_time_label.pack(side=tk.LEFT)

        self.bar_canvas = tk.Canvas(
            bar_frame, height=int(12 * s),
            bg="#222222", highlightthickness=0)
        self.bar_canvas.pack(side=tk.LEFT, fill=tk.X, expand=True,
                             padx=(int(10 * s), 0))

        self.question_label = tk.Label(
            self.root, text="",
            font=("Georgia", fs(36)),
            bg=bg, fg="white",
            wraplength=int(0.78 * self.root.winfo_screenwidth()),
            justify="center",
            padx=int(20 * s), pady=int(10 * s))
        self.question_label.pack(pady=(int(20 * s), int(8 * s)), expand=True)

        self.hint_label = tk.Label(
            self.root,
            text="Jot down notes if you need!",
            bg=bg, fg="#555555",
            font=("Georgia", fs(13), "italic"))
        self.hint_label.pack()

        self.feedback_label = tk.Label(
            self.root, text="",
            font=("Courier New", fs(22), "bold"),
            bg=bg, fg="white")
        self.feedback_label.pack(pady=int(6 * s))

        self.btn_bg      = "#1a1a2e"
        self.btn_fg      = "#e0e0e0"
        self.btn_hover   = "#16213e"
        self.btn_correct = "#2d6a4f"
        self.btn_wrong   = "#9b2226"

        self.buttons_frame = tk.Frame(self.root, bg=bg)
        self.option_buttons = []
        for i in range(4):
            btn = tk.Button(
                self.buttons_frame, text="",
                font=("Courier New", fs(24)),
                bg=self.btn_bg, fg=self.btn_fg,
                activebackground=self.btn_hover,
                activeforeground=self.btn_fg,
                relief=tk.FLAT, bd=0,
                padx=int(20 * s), pady=int(12 * s),
                width=14, justify="center",
                command=lambda idx=i: self.check_answer(idx))
            self.option_buttons.append(btn)

    # ------------------------------------------------------------------
    def on_closing(self):
        self.running = False
        self.stop_music()
        self.tracker.stop()
        for tid in [self.current_timer_id, self.q_timer_id, self._hide_timer_id]:
            if tid:
                try:
                    self.root.after_cancel(tid)
                except Exception:
                    pass
        time.sleep(0.1)
        try:
            self.root.destroy()
        except tk.TclError:
            pass
        if self.callback:
            self.callback()

    # ------------------------------------------------------------------
    def countdown(self, count):
        if not self.running:
            return
        if count > 0:
            try:
                self.question_label.config(
                    text=str(count),
                    font=("Courier New", self._fs(110), "bold"))
                self.root.after(1000, self.countdown, count - 1)
            except tk.TclError:
                pass
        else:
            try:
                self.question_label.config(
                    text="", font=("Georgia", self._fs(36)))
                self.buttons_frame.pack(pady=int(14 * self.scale))
                for i, btn in enumerate(self.option_buttons):
                    row, col = divmod(i, 2)
                    btn.grid(row=row, column=col,
                             padx=int(12 * self.scale),
                             pady=int(8 * self.scale))
                self._start_music_thread()
                self.generate_question()
                self.update_timer()
            except tk.TclError:
                pass

    # ------------------------------------------------------------------
    def _play_music(self):
        path = get_resource_path("clock.wav")
        self.music_playing = True
        if _SYSTEM == "Windows":
            try:
                import winsound
                # Purge any leftover audio (e.g. relax.wav) before starting clock
                winsound.PlaySound(None, winsound.SND_PURGE)
                while self.music_playing:
                    winsound.PlaySound(path, winsound.SND_FILENAME)
            except Exception as exc:
                print("[Audio] {}".format(exc))
        else:
            # macOS / Linux: loop clock.wav with subprocess
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
                    break   # no player found

    def stop_music(self):
        self.music_playing = False
        if _SYSTEM == "Windows":
            try:
                import winsound
                winsound.PlaySound(None, winsound.SND_PURGE)
            except Exception:
                pass

    def _start_music_thread(self):
        threading.Thread(target=self._play_music, daemon=True).start()

    # ------------------------------------------------------------------
    def update_timer(self):
        if not self.running:
            return
        if self.remaining_time > 0:
            try:
                self.remaining_time -= 1
                self.time_label.config(
                    text="Time: {}".format(self.remaining_time))
                self.current_timer_id = self.root.after(1000, self.update_timer)
            except tk.TclError:
                pass
        else:
            self.end_game()

    def _start_q_bar(self, total_secs):
        self._q_bar_total   = total_secs
        self._q_bar_elapsed = 0
        self._update_q_bar()

    def _update_q_bar(self):
        if not self.running or self.answered:
            return
        try:
            elapsed  = self._q_bar_elapsed
            total    = self._q_bar_total
            fraction = max(0.0, 1.0 - elapsed / total)
            self.q_time_label.config(text="{}s".format(total - elapsed))
            w = self.bar_canvas.winfo_width()
            h = self.bar_canvas.winfo_height()
            self.bar_canvas.delete("all")
            colour = ("#55b87a" if fraction > 0.4 else
                      "#ffcc44" if fraction > 0.2 else "#d9534f")
            self.bar_canvas.create_rectangle(
                0, 0, int(w * fraction), h, fill=colour, outline="")
            self._q_bar_elapsed += 1
            self.q_timer_id = self.root.after(1000, self._update_q_bar)
        except tk.TclError:
            pass

    # ------------------------------------------------------------------
    def generate_question(self):
        if not self.running or self.remaining_time <= 0:
            return
        try:
            self.tracker.begin_question()
            self.question_start_time = time.time()
            self.current_q_time      = QUESTION_TIME[self.tier]

            q_text, correct, choices = QuestionEngine.generate(self.tier, self.seen_questions)
            self.question_text      = q_text
            self.correct_answer     = correct
            self.choices            = choices
            self.correct_option_idx = choices.index(correct)

            self.question_label.config(
                text=q_text, font=("Georgia", self._fs(36)))
            self.feedback_label.config(text="")
            self._refresh_score_label()

            letters = ["A", "B", "C", "D"]
            for i, btn in enumerate(self.option_buttons):
                btn.config(
                    text="{}:  {}".format(letters[i], choices[i]),
                    bg=self.btn_bg, fg=self.btn_fg, state=tk.NORMAL)

            self.answered = False
            for tid in [self.q_timer_id, self._hide_timer_id]:
                if tid:
                    self.root.after_cancel(tid)
            self._start_q_bar(self.current_q_time)
            self._hide_timer_id = self.root.after(
                self.current_q_time * 1000, self.hide_question)
        except tk.TclError:
            pass

    # ------------------------------------------------------------------
    def _promote(self):
        self.streak_correct += 1
        self.streak_wrong    = 0
        if self.streak_correct >= PROMOTE_THRESHOLD and self.tier < 2:
            self.tier += 1
            self.streak_correct = 0

    def _demote(self):
        self.streak_wrong   += 1
        self.streak_correct  = 0
        if self.streak_wrong >= DEMOTE_THRESHOLD and self.tier > 1:
            self.tier -= 1
            self.streak_wrong = 0

    def _refresh_tier_label(self):
        text, bg, fg = self.TIER_INFO[self.tier]
        self.tier_label.config(text=text, bg=bg, fg=fg)

    def _refresh_score_label(self):
        self.score_label.config(
            text="correct: {}   wrong: {}".format(
                self.correct_count, self.wrong_count))

    # ------------------------------------------------------------------
    def check_answer(self, idx):
        if not self.running or self.answered:
            return
        self.answered   = True
        response_ms     = int((time.time() - self.question_start_time) * 1000)
        movement        = self.tracker.end_question()

        for tid in [self._hide_timer_id, self.q_timer_id]:
            if tid:
                try:
                    self.root.after_cancel(tid)
                except Exception:
                    pass

        try:
            correct = (idx == self.correct_option_idx)
            self.total_questions += 1

            if correct:
                self.correct_count += 1
                self.option_buttons[idx].config(bg=self.btn_correct, fg="white")
                self.feedback_label.config(text="Correct!", fg="#55b87a")
                self.write_result_to_xlsx("CORRECT", idx, response_ms, movement)
                self._promote()
            else:
                self.wrong_count += 1
                self.option_buttons[idx].config(bg=self.btn_wrong, fg="white")
                self.option_buttons[self.correct_option_idx].config(
                    bg=self.btn_correct, fg="white")
                self.feedback_label.config(
                    text="Incorrect - answer was {}".format(self.correct_answer),
                    fg="#ff6b6b")
                self.write_result_to_xlsx("WRONG", idx, response_ms, movement)
                self._demote()

            self._refresh_score_label()
            for btn in self.option_buttons:
                btn.config(state=tk.DISABLED)
            self.q_time_label.config(text="")
            self.root.after(1000, self._advance)
        except tk.TclError:
            pass

    def _advance(self):
        if not self.running:
            return
        try:
            for btn in self.option_buttons:
                btn.config(text="", bg=self.btn_bg, state=tk.NORMAL)
            self.feedback_label.config(text="")
            self.generate_question()
        except tk.TclError:
            pass

    def hide_question(self):
        if not self.running or self.answered:
            return
        self.answered   = True
        response_ms     = int((time.time() - self.question_start_time) * 1000)
        movement        = self.tracker.end_question()
        self.total_questions += 1
        self.miss_count      += 1
        self.write_result_to_xlsx("MISS", -1, response_ms, movement)
        self._demote()
        try:
            self.feedback_label.config(
                text="Time's up!  Answer: {}".format(self.correct_answer),
                fg="#ffcc44")
            self.q_time_label.config(text="")
            self._refresh_score_label()
            for btn in self.option_buttons:
                btn.config(text="", bg=self.btn_bg)
            self.root.after(1000, self._advance)
        except tk.TclError:
            pass

    # ------------------------------------------------------------------
    def end_game(self):
        if not self.running:
            return
        for tid in [self._hide_timer_id, self.q_timer_id, self.current_timer_id]:
            if tid:
                try:
                    self.root.after_cancel(tid)
                except Exception:
                    pass
        self.running  = False
        self.answered = True
        self.tracker.stop()
        try:
            self.question_label.config(
                text="Game Over!",
                font=("Courier New", self._fs(80), "bold"))
            self.feedback_label.config(text="")
            self.hint_label.config(text="")
            self.q_time_label.config(text="")
            for btn in self.option_buttons:
                btn.config(text="", state=tk.DISABLED)
        except tk.TclError:
            pass
        append_timeline_event(self.path_to_file, "Test End")
        self.write_summary_to_xlsx()
        self.stop_music()
        try:
            self.root.after(2500, self.root.destroy)
        except tk.TclError:
            pass
        if self.callback:
            self.callback()

    # ------------------------------------------------------------------
    def write_result_to_xlsx(self, result, idx, response_ms, movement):
        selected = self.option_buttons[idx].cget("text") if idx >= 0 else "MISS"
        row = [
            time.strftime('%m/%d/%Y %H:%M:%S', time.localtime()),
            int((time.time() - self.start_time) * 1000),
            response_ms, self.tier, self.current_q_time,
            self.question_text.replace('\n', ' '),
            self.correct_answer, selected, result,
            movement["sample_count"],
            movement["mean_magnitude"],
            movement["max_magnitude"],
        ]
        threading.Thread(
            target=append_trial_to_xlsx,
            args=(self.path_to_file, row),
            daemon=True).start()

    def write_summary_to_xlsx(self):
        summary = {
            "Start Time":   self.start_time_header,
            "Total Q":      self.total_questions,
            "Correct":      self.correct_count,
            "Wrong":        self.wrong_count,
            "Miss":         self.miss_count,
            "Final Tier":   self.tier,
            "Test Version": TEST_LABEL,
        }
        write_summary_to_xlsx(self.path_to_file, summary)


# ---------------------------------------------------------------------------
# Thank You Screen  (shown after post-test relaxation)
# ---------------------------------------------------------------------------
class ThankYouScreen:
    """
    Closing screen shown after the post-test relaxation period ends.
    Matches the dark calming aesthetic of the rest of the session.
    """

    BG        = "#0a0e1a"
    CARD_BG   = "#111828"
    BORDER    = "#1e3050"
    TEXT_HEAD = "#c8d8f0"
    TEXT_BODY = "#8aa8cc"
    TEXT_DIM  = "#4a6080"

    def __init__(self, master):
        self.master = master
        self.master.title("Session Complete")
        self.master.configure(bg=self.BG)
        self.master.protocol("WM_DELETE_WINDOW", self.master.destroy)

        sw = master.winfo_screenwidth()
        sh = master.winfo_screenheight()
        self._base = min(sw, sh)
        ww = max(520, min(int(sw * 0.50), 780))
        wh = max(380, min(int(sh * 0.55), 560))
        x  = (sw - ww) // 2
        y  = (sh - wh) // 2
        master.geometry(f"{ww}x{wh}+{x}+{y}")
        master.resizable(True, True)
        self._ww = ww
        self._wh = wh
        self._build_ui()

    def _fs(self, pt):
        return max(8, int(pt * self._base / 1080))

    def _build_ui(self):
        outer = tk.Frame(self.master, bg=self.BG)
        outer.pack(fill=tk.BOTH, expand=True)

        card_wrap = tk.Frame(outer, bg=self.BORDER)
        card_wrap.place(relx=0.5, rely=0.5, anchor="center",
                        relwidth=0.88, relheight=0.88)

        card = tk.Frame(card_wrap, bg=self.CARD_BG)
        card.pack(fill=tk.BOTH, expand=True, padx=1, pady=1)

        inner = tk.Frame(card, bg=self.CARD_BG)
        inner.pack(fill=tk.BOTH, expand=True,
                   padx=int(self._ww * 0.08),
                   pady=int(self._wh * 0.08))

        # Checkmark icon (unicode)
        tk.Label(
            inner, text="✓",
            font=("Helvetica Neue", self._fs(52)),
            bg=self.CARD_BG, fg="#4a90d9"
        ).pack(pady=(0, int(self._wh * 0.02)))

        tk.Label(
            inner,
            text="Thank You",
            font=("Helvetica Neue", self._fs(28), "bold"),
            bg=self.CARD_BG, fg=self.TEXT_HEAD
        ).pack()

        tk.Frame(inner, bg=self.BORDER, height=1).pack(
            fill=tk.X, pady=(int(self._wh * 0.03), int(self._wh * 0.035)))

        tk.Label(
            inner,
            text="You have completed the practice test.\n\nYou will be directed to take the harder version next.\nPlease let the researcher know you are ready.",
            font=("Helvetica Neue", self._fs(14)),
            bg=self.CARD_BG, fg=self.TEXT_BODY,
            justify="center",
            wraplength=int(self._ww * 0.70)
        ).pack()

        # Spacer
        tk.Frame(inner, bg=self.CARD_BG).pack(fill=tk.BOTH, expand=True)

        tk.Label(
            inner,
            text="You may close this window.",
            font=("Helvetica Neue", self._fs(11), "italic"),
            bg=self.CARD_BG, fg=self.TEXT_DIM
        ).pack(pady=(0, int(self._wh * 0.01)))


# ---------------------------------------------------------------------------
# Session orchestrator
# ---------------------------------------------------------------------------
# Each stage owns its own tk.Tk() + mainloop().
# on_done / on_ready / callback are called AFTER the window is destroyed,
# so they simply call the next stage function which spins up a fresh root.
# ---------------------------------------------------------------------------

def run_pre_relaxation(participant_data):
    """Step 2: pre-test 3-minute relaxation.
    Excel file is created here so relaxation timestamps can be recorded.
    """
    # Create the Excel file now (before the test) so we can log relax times
    xlsx = get_xlsx_path(participant_data["username"])
    create_xlsx_skeleton(xlsx, participant_data)
    # Store path in participant_data so MathTestUI picks it up
    participant_data["_xlsx_path"] = xlsx

    root = tk.Tk()
    RelaxationScreen(
        root,
        duration=RELAX_DURATION,
        on_done=lambda: show_instructions(participant_data),
        xlsx_path=xlsx,
        period_label="Pre-Test Relaxation",
    )
    root.mainloop()


def show_instructions(participant_data):
    """Step 3: wake-up audio + instructions card with Start button."""
    root = tk.Tk()
    InstructionsWindow(
        root,
        on_ready=lambda: launch_test(participant_data),
    )
    root.mainloop()


def launch_test(participant_data):
    """Step 4: math test."""
    root = tk.Tk()
    sw = root.winfo_screenwidth()
    sh = root.winfo_screenheight()
    ww = max(900,  min(1920, int(sw * 0.90)))
    wh = max(650,  min(1080, int(sh * 0.90)))
    root.geometry("{}x{}+{}+{}".format(ww, wh, (sw - ww) // 2, (sh - wh) // 2))
    root.minsize(900, 650)
    MathTestUI(
        root,
        username=participant_data["username"],
        participant_data=participant_data,
        callback=lambda: run_post_relaxation(participant_data),
    )
    root.mainloop()


def run_post_relaxation(participant_data):
    """Step 5: post-test 3-minute relaxation."""
    xlsx = participant_data.get("_xlsx_path")
    root = tk.Tk()
    RelaxationScreen(
        root,
        duration=RELAX_DURATION,
        on_done=show_thank_you,
        xlsx_path=xlsx,
        period_label="Post-Test Relaxation",
    )
    root.mainloop()


def show_thank_you():
    """Step 6: closing thank-you screen."""
    root = tk.Tk()
    ThankYouScreen(root)
    root.mainloop()


# ---------------------------------------------------------------------------
# Entry point
# ---------------------------------------------------------------------------
if __name__ == "__main__":
    intake_root = tk.Tk()
    ParticipantIntakeUI(intake_root, on_submit=run_pre_relaxation)
    intake_root.mainloop()