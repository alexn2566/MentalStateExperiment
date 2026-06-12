"""
Math Test — Version B: Hard (Scratch Work Required)
=====================================================
All questions require the user to write down working steps.
No question should be doable purely in your head, but none
requires a calculator — clean integer or simple fraction answers.

Fixed 35 seconds per question throughout. No tier adaptation —
difficulty stays constant to isolate the scratch-work effect.

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
import random
import csv
import time
import threading
import os
import sys
import math
import fractions

# ---------------------------------------------------------------------------
def get_resource_path(relative_path: str) -> str:
    base = getattr(sys, '_MEIPASS', os.path.dirname(os.path.abspath(__file__)))
    return os.path.join(base, relative_path)

# ---------------------------------------------------------------------------
TOTAL_GAME_TIME   = 60 * 3
COUNTDOWN         = 5
PATH              = './data/'
QUESTION_TIME_SEC = 35          # fixed — all questions get 35 s
NICLA_ENABLED     = True
NICLA_CHAR_UUID   = "19b10001-e8f2-537e-4f6c-d104768a1214"
NICLA_DEVICE_NAME = "NiclaSenseME"
TEST_LABEL        = "Hard — Scratch Work"
CSV_PREFIX        = "mathTest_hard_"

# ---------------------------------------------------------------------------
# Hand-movement tracker
# ---------------------------------------------------------------------------
class HandMovementTracker:
    def __init__(self):
        self.enabled  = False
        self._lock    = threading.Lock()
        self._samples: list[float] = []
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

    # ------------------------------------------------------------------
    @staticmethod
    def _simultaneous() -> tuple[str, str]:
        """
        Two simultaneous linear equations.
        ax + by = p
        cx + dy = q
        Solve for x (integer solution guaranteed).
        """
        # Build from solution outward so answer is always a clean integer
        x = random.randint(2, 10)
        y = random.randint(2, 10)
        a = random.randint(2, 6)
        b = random.randint(1, 5)
        c = random.randint(1, 5)
        d = random.randint(2, 6)
        # Ensure unique solution: ad - bc ≠ 0
        while a * d - b * c == 0:
            d = random.randint(2, 6)
        p = a * x + b * y
        q = c * x + d * y
        q_text = (f"Solve the system:\n"
                  f"  {a}x + {b}y = {p}\n"
                  f"  {c}x + {d}y = {q}\n"
                  f"What is the value of x?")
        return q_text, str(x)

    # ------------------------------------------------------------------
    @staticmethod
    def _quadratic() -> tuple[str, str]:
        """
        x² + bx + c = 0  where roots are integers.
        Ask for the positive root (or larger root).
        """
        r1 = random.randint(-8, -1)
        r2 = random.randint(1, 8)
        # (x - r1)(x - r2) = x² - (r1+r2)x + r1*r2
        b = -(r1 + r2)
        c = r1 * r2
        b_str = f"+ {b}" if b >= 0 else f"- {abs(b)}"
        c_str = f"+ {c}" if c >= 0 else f"- {abs(c)}"
        larger_root = max(r1, r2)
        q_text = (f"Solve:  x² {b_str}x {c_str} = 0\n"
                  f"What is the larger root?")
        return q_text, str(larger_root)

    # ------------------------------------------------------------------
    @staticmethod
    def _geometry() -> tuple[str, str]:
        """
        Geometry requiring multiple formula steps.
        Subtypes: composite area, similar triangles, Pythagorean triple check,
                  perimeter from area.
        """
        sub = random.choice(['composite', 'similar', 'pyth', 'rect_from_area'])

        if sub == 'composite':
            # L-shaped figure: big rectangle minus small rectangle
            W  = random.randint(8, 16)
            H  = random.randint(8, 16)
            w  = random.randint(2, W - 2)
            h  = random.randint(2, H - 2)
            ans = W * H - w * h
            q_text = (f"An L-shaped figure is made by taking a {W}×{H} rectangle "
                      f"and removing a {w}×{h} rectangle from one corner.\n"
                      f"What is the area of the L-shape?")

        elif sub == 'similar':
            # Two similar triangles; find missing side
            scale = random.randint(2, 4)
            a1    = random.randint(3, 10)
            b1    = random.randint(3, 10)
            a2    = a1 * scale
            b2    = b1 * scale
            # Give a1, b1, a2; find b2
            ans   = b2
            q_text = (f"Two similar triangles have corresponding sides in proportion.\n"
                      f"Triangle 1 has sides {a1} and {b1}.\n"
                      f"Triangle 2 has a side of {a2} corresponding to the {a1} side.\n"
                      f"What is the side of Triangle 2 that corresponds to {b1}?")

        elif sub == 'pyth':
            # Give two legs, find hypotenuse (Pythagorean triple)
            triples = [(3,4,5),(5,12,13),(8,15,17),(7,24,25),(9,40,41),(6,8,10),(9,12,15)]
            a, b, c = random.choice(triples)
            mult    = random.randint(1, 3)
            a, b, c = a*mult, b*mult, c*mult
            ans     = c
            q_text  = (f"A right triangle has legs of length {a} and {b}.\n"
                       f"What is the length of the hypotenuse?")

        else:  # rect_from_area
            # Given area and one side relation, find perimeter
            # width = w, length = w + k, area = w(w+k)
            w = random.randint(4, 12)
            k = random.randint(2, 8)
            area = w * (w + k)
            perim = 2 * (w + w + k)
            ans   = perim
            q_text = (f"A rectangle has an area of {area} square units.\n"
                      f"Its length is {k} more than its width.\n"
                      f"What is its perimeter?")

        return q_text, str(ans)

    # ------------------------------------------------------------------
    @staticmethod
    def _sequence_sum() -> tuple[str, str]:
        """
        Sum of first N terms of an arithmetic sequence.
        S = N/2 * (2a + (N-1)d)
        """
        a = random.randint(1, 10)    # first term
        d = random.randint(1, 8)     # common difference
        n = random.randint(5, 12)    # number of terms
        S = n * (2 * a + (n - 1) * d) // 2
        q_text = (f"Find the sum of the first {n} terms of the arithmetic sequence "
                  f"that starts at {a} with a common difference of {d}.")
        return q_text, str(S)

    # ------------------------------------------------------------------
    @staticmethod
    def _pct_chain() -> tuple[str, str]:
        """
        Two successive percentage changes applied to a starting value.
        e.g. price goes up 20% then down 15% — what is the final value?
        Numbers chosen so the result is a clean integer.
        """
        # Pick starting value and two percentage changes that give integer result
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

    # ------------------------------------------------------------------
    @staticmethod
    def _age_problem() -> tuple[str, str]:
        """
        Classic two-person age word problem requiring algebra.
        """
        sub = random.choice(['ratio_now', 'future_sum', 'diff_and_sum'])

        if sub == 'ratio_now':
            # Ages now in ratio a:b, sum = s
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
            # Now A=a, B=b; in t years their ages sum to s
            a = random.randint(10, 30)
            b = random.randint(10, 30)
            t = random.randint(3, 10)
            s = (a + t) + (b + t)
            q_text = (f"Person A is currently {a} years old and Person B is {b}.\n"
                      f"In how many years will the sum of their ages be {s}?")
            ans = t

        else:  # diff and sum
            diff = random.randint(2, 15)
            total = random.randint(20, 60)
            # older = (total + diff) / 2 — ensure integer
            while (total + diff) % 2 != 0:
                total += 1
            older = (total + diff) // 2
            q_text = (f"The sum of two people's ages is {total}. "
                      f"One person is {diff} years older than the other.\n"
                      f"How old is the older person?")
            ans = older

        return q_text, str(ans)

    # ------------------------------------------------------------------
    @staticmethod
    def _mixture() -> tuple[str, str]:
        """
        Mixture problem: combine two solutions of different concentrations.
        How many litres of solution A (c1%) must be mixed with solution B (c2%)
        to get V litres of target concentration ct%?
        Amount of A = V*(ct - c2)/(c1 - c2)
        Choose values to give a clean integer answer.
        """
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
        # Fallback if no clean integer found quickly
        return HardQuestionEngine._sequence_sum()

    # ------------------------------------------------------------------
    @staticmethod
    def _work_three() -> tuple[str, str]:
        """
        Three workers; A alone in a days, B alone in b days, C alone in c days.
        Working together, how many days? (fractional answer allowed)
        """
        # Choose a, b, c so that 1/a+1/b+1/c gives a simple fraction
        combos = [
            (2, 3, 6),   # sum = 1  → 1 day
            (3, 4, 6),   # sum = 3/4 → 4/3 days
            (2, 4, 4),   # sum = 1  → 1 day
            (3, 6, 6),   # sum = 2/3 → 3/2 days
            (4, 6, 12),  # sum = 1/2 → 2 days
            (2, 6, 3),
            (6, 4, 12),
            (3, 3, 6),
        ]
        a, b, c = random.choice(combos)
        total_rate = fractions.Fraction(1, a) + fractions.Fraction(1, b) + fractions.Fraction(1, c)
        ans = str(fractions.Fraction(1, 1) / total_rate)
        q_text = (f"Worker A completes a job alone in {a} days, "
                  f"Worker B in {b} days, and Worker C in {c} days.\n"
                  f"How many days does it take all three working together?\n"
                  f"(Express as a fraction if needed)")
        return q_text, ans

    # ------------------------------------------------------------------
    @staticmethod
    def _digit_problem() -> tuple[str, str]:
        """
        Number theory / digit problem requiring algebraic setup.
        """
        sub = random.choice(['reverse', 'sum_and_diff', 'divisibility'])

        if sub == 'reverse':
            # tens digit t, units digit u; number = 10t+u
            # reversed = 10u+t; reversed - original = k
            t = random.randint(2, 7)
            u = random.randint(t + 1, 9)  # u > t so reversed > original
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

        else:  # divisibility
            base = random.randint(2, 8)
            mult = random.randint(3, 9)
            target = base * mult
            # "What is the smallest two-digit multiple of N greater than M?"
            M     = random.randint(10, 60)
            start = (M // target + 1) * target
            q_text = (f"What is the smallest two-digit multiple of {target} "
                      f"that is greater than {M}?")
            ans = start

        return q_text, str(ans)

    # ------------------------------------------------------------------
    @staticmethod
    def _proportion_word() -> tuple[str, str]:
        """
        Multi-step ratio / proportion problem.
        """
        sub = random.choice(['map_scale', 'recipe_scale', 'gear_ratio'])

        if sub == 'map_scale':
            scale_cm  = random.randint(2, 5)    # cm on map
            scale_km  = random.randint(10, 50)  # km in real life
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

        else:  # gear_ratio
            t1 = random.choice([12, 15, 18, 20, 24, 30])
            t2 = random.choice([6,  8,  9,  10, 12, 15])
            rpm1 = random.choice([60, 80, 100, 120, 150])
            # t1*rpm1 = t2*rpm2
            rpm2 = t1 * rpm1 // t2
            q_text = (f"Gear A has {t1} teeth and spins at {rpm1} rpm. "
                      f"It meshes with Gear B which has {t2} teeth.\n"
                      f"How fast does Gear B spin (in rpm)?")
            ans = rpm2

        return q_text, str(ans)

    # ------------------------------------------------------------------
    # Public interface
    # ------------------------------------------------------------------
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
    def generate(cls) -> tuple[str, str, list[str]]:
        q_text, correct = random.choice(cls._GENERATORS)()
        return cls._pack(q_text, correct)

    @staticmethod
    def _pack(question: str, correct: str) -> tuple[str, str, list[str]]:
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
        self.path_to_file        = PATH + CSV_PREFIX + username + '.csv'
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
        self.question_text       = ""
        self.correct_answer      = ""
        self.choices: list[str]  = []
        self.correct_option_idx  = 0
        self.music_playing       = False

        self.tracker = HandMovementTracker()
        self.tracker.start()

        self._build_ui(fs)

        self.start_time        = time.time()
        self.start_time_header = time.strftime('%m/%d/%Y %H:%M:%S',
                                               time.localtime(self.start_time))
        os.makedirs(os.path.dirname(os.path.abspath(self.path_to_file)),
                    exist_ok=True)
        print("Saving data to:", os.path.abspath(self.path_to_file))
        self.write_header_to_csv()
        self.countdown(COUNTDOWN)

    # ------------------------------------------------------------------ UI
    def _build_ui(self, fs):
        BG = "#0d0d0d"
        s  = self.scale

        top = tk.Frame(self.root, bg=BG)
        top.pack(fill=tk.X, padx=int(18*s), pady=int(12*s))

        self.time_label = tk.Label(top, text=f"Time: {self.remaining_time}",
                                   bg=BG, fg="white",
                                   font=("Courier New", fs(20), "bold"),
                                   anchor='w')
        self.time_label.pack(side=tk.LEFT)

        self.tier_label = tk.Label(top, text="Hard — Scratch Work Required",
                                   bg="#5a1a1a", fg="#ff6b6b",
                                   font=("Courier New", fs(16), "bold"),
                                   padx=int(12*s), pady=int(4*s))
        self.tier_label.pack(side=tk.LEFT, padx=int(20*s))

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

        # Question box with slightly larger font — questions are longer
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
                self.update_timer()
            except tk.TclError: pass

    # ------------------------------------------------------------------ audio
    def _play_music(self):
        import winsound
        path = get_resource_path("clock.wav")
        self.music_playing = True
        try:
            while self.music_playing:
                winsound.PlaySound(path, winsound.SND_FILENAME)
        except Exception as e:
            print(f"[Audio] {e}")

    def stop_music(self):
        self.music_playing = False
        try:
            import winsound
            winsound.PlaySound(None, winsound.SND_PURGE)
        except Exception: pass

    def _start_music_thread(self):
        threading.Thread(target=self._play_music, daemon=True).start()

    # ------------------------------------------------------------------ timers
    def update_timer(self):
        if not self.running: return
        if self.remaining_time > 0:
            try:
                self.remaining_time -= 1
                self.time_label.config(text=f"Time: {self.remaining_time}")
                self.current_timer_id = self.root.after(1000, self.update_timer)
            except tk.TclError: pass
        else:
            self.end_game()

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
        if not self.running or self.remaining_time <= 0: return
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
                self.write_result_to_csv("CORRECT", idx, response_ms, movement)
            else:
                self.wrong_count += 1
                self.option_buttons[idx].config(bg=self.btn_wrong, fg="white")
                self.option_buttons[self.correct_option_idx].config(
                    bg=self.btn_correct, fg="white")
                self.feedback_label.config(
                    text=f"✗  Incorrect — answer was {self.correct_answer}",
                    fg="#ff6b6b")
                self.write_result_to_csv("WRONG", idx, response_ms, movement)
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
        self.write_result_to_csv("MISS", -1, response_ms, movement)
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
                text="Game Over!",
                font=("Courier New", self._fs(80), "bold"))
            self.feedback_label.config(text="")
            self.hint_label.config(text="")
            self.q_time_label.config(text="")
            for btn in self.option_buttons:
                btn.config(text="", state=tk.DISABLED)
        except tk.TclError: pass
        self.write_summary_to_csv()
        self.stop_music()
        try: self.root.after(2500, self.root.destroy)
        except tk.TclError: pass
        if self.callback: self.callback()

    # ------------------------------------------------------------------ CSV
    def write_header_to_csv(self):
        pd = self.participant_data
        with open(self.path_to_file, 'a', newline='') as f:
            w = csv.writer(f)
            w.writerow(["Participant", self.username])
            w.writerow(["Email",  pd.get("email", "")])
            w.writerow(["Age",    pd.get("age", "")])
            w.writerow(["Gender", pd.get("gender", "")])
            w.writerow([])
            w.writerow([
                "Timestamp", "Relative Time (ms)", "Response Time (ms)",
                "Q Time Allowed (s)", "Question",
                "Correct Answer", "Selected Option", "Correctness",
                "HM_Samples", "HM_Mean_Mag", "HM_Max_Mag", self.username])

    def write_result_to_csv(self, result, idx, response_ms, movement):
        selected = self.option_buttons[idx].cget("text") if idx >= 0 else "MISS"
        with open(self.path_to_file, 'a', newline='') as f:
            csv.writer(f).writerow([
                time.strftime('%m/%d/%Y %H:%M:%S', time.localtime()),
                int((time.time() - self.start_time) * 1000),
                response_ms, QUESTION_TIME_SEC,
                self.question_text.replace('\n', ' '),
                self.correct_answer, selected, result,
                movement["sample_count"], movement["mean_magnitude"],
                movement["max_magnitude"]])

    def write_summary_to_csv(self):
        with open(self.path_to_file, 'r', newline='') as rf:
            lines = list(csv.reader(rf))
        hdr = ["Start Time", "Total Q", "Correct", "Wrong", "Miss",
               "Test Version", self.username]
        row = [self.start_time_header, self.total_questions,
               self.correct_count, self.wrong_count, self.miss_count,
               TEST_LABEL]
        lines.insert(0, hdr); lines.insert(1, row)
        with open(self.path_to_file, 'w', newline='') as wf:
            w = csv.writer(wf)
            w.writerows(lines)
            w.writerow(hdr); w.writerow(row)



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

    # ------------------------------------------------------------------
    def _build_ui(self):
        # Set up ttk style FIRST before any widgets are created
        from tkinter import ttk
        style = ttk.Style()
        style.theme_use("clam")
        style.configure("Intake.TCombobox",
                         fieldbackground=IN_ENTRY_BG,
                         background=IN_ENTRY_BG,
                         foreground=IN_TEXT,
                         bordercolor=IN_BORDER,
                         arrowcolor=IN_ACCENT,
                         padding=6)

        # Scrollable canvas wrapper so the button is always reachable
        canvas = tk.Canvas(self.root, bg=IN_BG, highlightthickness=0)
        scrollbar = tk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        canvas.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scroll_frame = tk.Frame(canvas, bg=IN_BG)
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

        # First / Last name row
        row = tk.Frame(outer, bg=IN_BG)
        row.pack(fill=tk.X, pady=(0, 12))
        self._field(row, "first_name", "First name *", side=tk.LEFT, width=18)
        tk.Frame(row, bg=IN_BG, width=16).pack(side=tk.LEFT)
        self._field(row, "last_name",  "Last name *",  side=tk.LEFT, width=18)

        self._field(outer, "email", "Email address *", full=True)
        self._field(outer, "age",   "Age *",           full=True, numeric=True)

        # Gender dropdown
        tk.Label(outer, text="Gender", bg=IN_BG, fg=IN_MUTED,
                 font=("Helvetica Neue", 12)).pack(anchor="w", pady=(0, 4))
        gender_var = tk.StringVar(value=self.GENDERS[0])
        self._vars["gender"] = gender_var
        ttk.Combobox(outer, textvariable=gender_var,
                     values=self.GENDERS, state="readonly",
                     style="Intake.TCombobox",
                     font=("Helvetica Neue", 13)).pack(fill=tk.X, pady=(0, 16))

        # Consent box
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

        # Error label
        self._error_label = tk.Label(outer, text="", bg=IN_BG, fg=IN_ERROR,
                                      font=("Helvetica Neue", 11))
        self._error_label.pack(anchor="w", pady=(0, 6))

        # Start button — starts visually greyed out
        self._start_btn = tk.Button(
            outer, text="Start test →",
            bg="#aaaaaa", fg="#dddddd",
            font=("Helvetica Neue", 14, "bold"),
            activebackground="#16213e", activeforeground=IN_BTN_FG,
            relief=tk.FLAT, bd=0, padx=20, pady=12,
            state=tk.DISABLED,
            command=self._submit)
        self._start_btn.pack(fill=tk.X)

    # ------------------------------------------------------------------
    def _field(self, parent, key, label, side=None, width=None,
               full=False, numeric=False):
        from tkinter import ttk
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

    # ------------------------------------------------------------------
    def _toggle_button(self):
        if self._consent_var.get():
            self._start_btn.config(state=tk.NORMAL, bg=IN_ACCENT, fg=IN_BTN_FG)
        else:
            self._start_btn.config(state=tk.DISABLED, bg="#aaaaaa", fg="#dddddd")

    # ------------------------------------------------------------------
    def _validate(self):
        import re
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

    # ------------------------------------------------------------------
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
if __name__ == "__main__":
    def launch_test(participant_data):
        test_root = tk.Tk()
        sw, sh = test_root.winfo_screenwidth(), test_root.winfo_screenheight()
        ww = max(900, min(1920, int(sw * 0.90)))
        wh = max(650, min(1080, int(sh * 0.90)))
        test_root.geometry(f"{ww}x{wh}+{(sw-ww)//2}+{(sh-wh)//2}")
        test_root.minsize(900, 650)
        HardMathTest(test_root,
                     username=participant_data["username"],
                     participant_data=participant_data)
        test_root.mainloop()

    intake_root = tk.Tk()
    ParticipantIntakeUI(intake_root, on_submit=launch_test)
    intake_root.mainloop()