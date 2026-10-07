# -*- coding: utf-8 -*-
"""헬스일지 PC 창 — 서버 켜기 + AI 코치를 터미널 없이.
   실행: 바탕화면 '헬스일지' 바로가기 (pythonw panel.py, 콘솔 창 없음)

- 창을 열면 서버(serve.py)가 같은 프로세스의 스레드로 켜지고, 창을 닫으면 꺼진다.
- AI 코치: 날짜·부위·시간·컨디션을 넣고 [계획 만들기] → coach.make_plan() → plan.json → 폰 앱 홈 카드.
- print/서버 로그는 창 아래 '서버 로그'로 모은다 (pythonw 는 콘솔이 없어서 stdout 이 None).
"""
import datetime, json, os, queue, re, sys, threading, traceback
import tkinter as tk
from tkinter import ttk

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, HERE)

PARTS = ["자동", "하체", "등", "가슴", "어깨", "팔", "코어"]
CONDITIONS = ["감기 기운", "잠 부족", "피곤함", "근육통"]   # 체크하면 컨디션 메모 앞에 붙음
NOISE = re.compile(r'"(GET|HEAD) [^"]*" (200|304) ')   # 정적 파일 요청 로그는 숨김
FONT = ("맑은 고딕", 10)
FONT_B = ("맑은 고딕", 10, "bold")
FONT_H = ("맑은 고딕", 13, "bold")
ORANGE, GREEN, RED, DIM, BG = "#ff7a1a", "#2e9e55", "#d64545", "#777777", "#ffffff"


ERROR_LOG = os.path.join(HERE, "coach_log", "panel_error.log")   # coach_log/ 는 git 제외·서빙 차단


def write_error(text):
    try:
        os.makedirs(os.path.dirname(ERROR_LOG), exist_ok=True)
        with open(ERROR_LOG, "a", encoding="utf-8") as f:
            f.write(f"--- {datetime.datetime.now():%Y-%m-%d %H:%M:%S}\n{text}\n")
    except OSError:
        pass


class QueueWriter:
    """sys.stdout/stderr 대신 — 다른 스레드에서 쓴 글을 큐에 넣고 창이 꺼내 보여줌"""
    def __init__(self, q):
        self.q = q

    def write(self, s):
        if s:
            self.q.put(("log", s))

    def flush(self):
        pass


def format_plan(r):
    """make_plan 결과 → 창에 보여줄 글 [(문장, 태그)]"""
    p, m = r["plan"], r["meta"]
    out = [(p["title"] + "\n", "h"), (p["note"] + "\n\n", "dim")]
    for it in p["items"]:
        warm = f"W {it['warm']['kg']:g}×{it['warm']['reps']} · " if it.get("warm") else ""
        kg = "" if it.get("kg") is None else f"{it['kg']:g}kg × "
        reps = it.get("repsText") or it["reps"]
        unit = "" if str(reps).endswith("초") else "회"
        out.append((f"• {it['ex']}  ", "b"))
        out.append((f"{warm}{kg}{reps}{unit} · {it['sets']}세트\n", ""))
        if it.get("note"):
            out.append((f"   {it['note']}\n", "dim"))
    auto = [f"{rule} {msg}" for rule, msg in r["fixes"] if rule != "G6"]
    if auto:
        out.append(("\n자동수정\n", "b"))
        out += [(f"• {a}\n", "warn") for a in auto]
    llm_warn = [w for w in p.get("warnings", []) if not w.startswith(("[자동수정]", "[주의]"))]
    if llm_warn:
        out.append(("\n주의\n", "b"))
        out += [(f"• {w}\n", "dim") for w in llm_warn]
    saved = {"추가": "plan.json 에 추가", "교체": "plan.json 의 같은 날짜 계획을 교체"}.get(r["saved"], "미리보기 (저장 안 함)")
    out.append((f"\n{saved} · {m['model']} · {m['ms'] / 1000:.0f}초 · 약 ${m['cost_usd']}\n", "dim"))
    return out


class Panel(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("헬스일지")
        self.geometry("640x880")
        self.minsize(520, 600)
        try:
            self.iconbitmap(os.path.join(HERE, "icon.ico"))
        except tk.TclError:
            pass
        self.q = queue.Queue()
        sys.stdout = sys.stderr = QueueWriter(self.q)
        self.httpd = None
        self.busy = False
        self._build()
        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.after(100, self.poll)
        self.start_server()
        # 다른 창 뒤에 열리지 않게 잠깐 맨 앞으로
        self.lift()
        self.attributes("-topmost", True)
        self.after(500, lambda: self.attributes("-topmost", False))
        self.focus_force()
        if not (os.environ.get("ANTHROPIC_API_KEY") or os.environ.get("ANTHROPIC_AUTH_TOKEN")):
            self.set_status("API 키가 없어요 — 환경변수 ANTHROPIC_API_KEY 설정 후 창을 다시 여세요", RED)

    # ---------- 화면 ----------
    def _build(self):
        st = ttk.Style(self)
        st.theme_use("clam")
        self.configure(bg=BG)
        st.configure(".", font=FONT, background=BG)
        st.configure("TButton", padding=6, background="#f1f1f1", bordercolor="#cccccc")
        st.map("TButton", background=[("active", "#e6e6e6")])
        st.configure("Accent.TButton", font=FONT_B, foreground="white", background=ORANGE, bordercolor=ORANGE, padding=8)
        st.map("Accent.TButton", background=[("disabled", "#e0b89a"), ("active", "#ff8f3d")])
        st.configure("TEntry", fieldbackground="white")
        st.configure("TCombobox", fieldbackground="white")
        st.configure("TSpinbox", fieldbackground="white")
        st.configure("H.TLabel", font=FONT_H)
        pad = {"padx": 14}

        # 서버
        top = ttk.Frame(self)
        top.pack(fill="x", pady=(12, 4), **pad)
        self.dot = tk.Label(top, text="●", font=("맑은 고딕", 14), fg=DIM, bg=BG)
        self.dot.pack(side="left")
        self.srv_text = ttk.Label(top, text="서버 꺼짐", font=FONT_B)
        self.srv_text.pack(side="left", padx=(4, 0))
        self.srv_btn = ttk.Button(top, text="서버 끄기", command=self.toggle_server)
        self.srv_btn.pack(side="right")
        self.url = ttk.Label(self, text="", foreground=DIM)
        self.url.pack(anchor="w", **pad)

        ttk.Separator(self).pack(fill="x", pady=10, **pad)
        ttk.Label(self, text="AI 코치", style="H.TLabel").pack(anchor="w", **pad)

        form = ttk.Frame(self)
        form.pack(fill="x", pady=(6, 0), **pad)
        form.columnconfigure(1, weight=1)

        ttk.Label(form, text="날짜").grid(row=0, column=0, sticky="w", pady=4)
        d = ttk.Frame(form)
        d.grid(row=0, column=1, sticky="w")
        self.date = tk.StringVar(value=datetime.date.today().isoformat())
        ttk.Entry(d, textvariable=self.date, width=12).pack(side="left")
        ttk.Button(d, text="오늘", width=5, command=lambda: self.set_day(0)).pack(side="left", padx=(6, 0))
        ttk.Button(d, text="내일", width=5, command=lambda: self.set_day(1)).pack(side="left", padx=(4, 0))

        ttk.Label(form, text="부위").grid(row=1, column=0, sticky="w", pady=4)
        f = ttk.Frame(form)
        f.grid(row=1, column=1, sticky="w")
        self.focus_var = tk.StringVar(value="자동")
        ttk.Combobox(f, textvariable=self.focus_var, values=PARTS, state="readonly", width=8).pack(side="left")
        ttk.Label(f, text="시간(분)").pack(side="left", padx=(16, 6))
        self.time_var = tk.StringVar(value=str(self.default_time()))
        ttk.Spinbox(f, from_=20, to=120, increment=5, textvariable=self.time_var, width=5).pack(side="left")

        ttk.Label(form, text="컨디션").grid(row=2, column=0, sticky="nw", pady=4)
        self.note = tk.Text(form, height=3, font=FONT, wrap="word", relief="solid", borderwidth=1,
                            highlightthickness=0)
        self.note.grid(row=2, column=1, sticky="ew", pady=4)
        c = ttk.Frame(form)
        c.grid(row=3, column=1, sticky="w")
        self.cond_vars = {k: tk.BooleanVar() for k in CONDITIONS}
        for k, v in self.cond_vars.items():
            ttk.Checkbutton(c, text=k, variable=v).pack(side="left", padx=(0, 10))
        ttk.Label(form, text="감기·몸살이면 전 부위, '허리·통증·뻐근'이면 그 부위는 증량 안 함. 일정(예: 출근 전 급하게)도 적으면 반영",
                  foreground=DIM, font=("맑은 고딕", 9), wraplength=500).grid(row=4, column=1, sticky="w")

        ttk.Label(form, text="보류").grid(row=5, column=0, sticky="nw", pady=(8, 4))
        h = ttk.Frame(form)
        h.grid(row=5, column=1, sticky="ew", pady=(8, 4))
        h.columnconfigure(0, weight=1)
        self.hold_text = ttk.Label(h, text="", wraplength=440)
        self.hold_text.grid(row=0, column=0, sticky="w")
        ttk.Button(h, text="편집", width=5, command=self.edit_holds).grid(row=0, column=1, sticky="e")
        self.refresh_holds()

        b = ttk.Frame(self)
        b.pack(fill="x", pady=(10, 0), **pad)
        self.make_btn = ttk.Button(b, text="계획 만들기 → 폰으로", style="Accent.TButton",
                                   command=lambda: self.run_coach(dry_run=False))
        self.make_btn.pack(side="left")
        self.preview_btn = ttk.Button(b, text="미리보기 (저장 안 함)", command=lambda: self.run_coach(dry_run=True))
        self.preview_btn.pack(side="left", padx=(8, 0))

        s = ttk.Frame(self)
        s.pack(fill="x", pady=(10, 0), **pad)
        ttk.Label(s, text="저장된 AI 계획").pack(side="left")
        self.plan_var = tk.StringVar()
        self.plan_pick = ttk.Combobox(s, textvariable=self.plan_var, state="readonly", width=40)
        self.plan_pick.pack(side="left", padx=(8, 0), fill="x", expand=True)
        ttk.Button(s, text="삭제", width=5, command=self.delete_selected).pack(side="left", padx=(8, 0))

        self.status = ttk.Label(self, text="", foreground=DIM)
        self.status.pack(anchor="w", pady=(8, 0), **pad)

        self.result = tk.Text(self, height=16, font=FONT, wrap="word", relief="solid", borderwidth=1,
                              padx=10, pady=8, state="disabled")
        self.result.pack(fill="both", expand=True, pady=(6, 0), **pad)
        self.result.tag_configure("h", font=FONT_H)
        self.result.tag_configure("b", font=FONT_B)
        self.result.tag_configure("dim", foreground=DIM)
        self.result.tag_configure("warn", foreground=RED)

        ttk.Label(self, text="서버 로그", foreground=DIM).pack(anchor="w", pady=(10, 0), **pad)
        self.log = tk.Text(self, height=5, font=("Consolas", 9), wrap="word", relief="solid", borderwidth=1,
                           state="disabled", foreground="#444")
        self.log.pack(fill="x", pady=(2, 12), **pad)
        self.refresh_plans()                    # 상태줄이 생긴 뒤에 (읽기 실패를 거기 표시)

    def default_time(self):
        try:
            import coach
            return coach.load_profile()["time_budget_min"]
        except Exception:
            return 60

    def set_day(self, n):
        self.date.set((datetime.date.today() + datetime.timedelta(days=n)).isoformat())

    def set_status(self, text, color=DIM):
        self.status.configure(text=text, foreground=color)

    def show_result(self, parts):
        self.result.configure(state="normal")
        self.result.delete("1.0", "end")
        for text, tag in parts:
            self.result.insert("end", text, tag or ())
        self.result.configure(state="disabled")

    def append_log(self, s):
        lines = [l for l in s.splitlines(True) if not NOISE.search(l)]
        if not lines:
            return
        self.log.configure(state="normal")
        self.log.insert("end", "".join(lines))
        self.log.see("end")
        self.log.configure(state="disabled")

    # ---------- 보류 ----------
    def refresh_holds(self):
        import coach
        holds = coach.load_profile().get("hold") or []
        if not holds:
            return self.hold_text.configure(text="없음 — 진단 전까지 뺄 운동이 있으면 [편집]", foreground=DIM)
        lines = []
        for h in holds:
            what = " · ".join((h.get("exercises") or []) + [f"{p} 전체" for p in h.get("parts") or []])
            lines.append(f"{what}  ({h.get('reason') or '보류'})")
        self.hold_text.configure(text="\n".join(lines), foreground=RED)

    def edit_holds(self):
        """보류 운동 고르기 — 운동목록을 부위별 체크박스로. 저장하면 coach_profile.json 의 hold 를 통째로 바꿈"""
        import coach, sync_excel
        try:
            rows = [e for e in sync_excel.read_excel()["exercises"] if e["p"]]
        except Exception as e:
            return self.set_status(f"운동목록을 못 읽었어요: {e}", RED)
        holds = coach.load_profile().get("hold") or []
        held = coach.held_exercises(rows, {"hold": holds})
        reasons = [h.get("reason") for h in holds if h.get("reason")]

        win = tk.Toplevel(self)
        win.title("보류 운동")
        win.configure(bg=BG)
        win.transient(self)
        win.grab_set()
        ttk.Label(win, text="진단·회복 전까지 계획에서 뺄 운동", font=FONT_B).pack(anchor="w", padx=14, pady=(12, 2))
        ttk.Label(win, text="체크한 운동은 AI가 고를 수 있는 목록에서 빠지고, 들어와도 자동으로 지워짐(G8)",
                  foreground=DIM, font=("맑은 고딕", 9)).pack(anchor="w", padx=14)
        grid = ttk.Frame(win)
        grid.pack(fill="both", padx=14, pady=8)
        vars_ = {}
        by_part = {}
        for e in rows:
            by_part.setdefault(e["p"], []).append(e["n"])
        for col, (part, names) in enumerate(by_part.items()):
            ttk.Label(grid, text=part, font=FONT_B).grid(row=0, column=col, sticky="w", padx=(0, 16))
            for r, n in enumerate(names, start=1):
                vars_[n] = tk.BooleanVar(value=n in held)
                ttk.Checkbutton(grid, text=n, variable=vars_[n]).grid(row=r, column=col, sticky="w", padx=(0, 16))
        rf = ttk.Frame(win)
        rf.pack(fill="x", padx=14)
        ttk.Label(rf, text="이유").pack(side="left")
        reason = tk.StringVar(value=" / ".join(reasons))
        ttk.Entry(rf, textvariable=reason).pack(side="left", fill="x", expand=True, padx=(8, 0))

        def save():
            picked = [n for n, v in vars_.items() if v.get()]
            coach.save_holds([{"exercises": picked, "reason": reason.get().strip() or "보류"}] if picked else [])
            self.refresh_holds()
            self.set_status(f"보류 {len(picked)}개 저장 — 다음 계획부터 반영", GREEN)
            win.destroy()
        bf = ttk.Frame(win)
        bf.pack(fill="x", padx=14, pady=12)
        ttk.Button(bf, text="저장", style="Accent.TButton", command=save).pack(side="right")
        ttk.Button(bf, text="취소", command=win.destroy).pack(side="right", padx=(0, 8))

    # ---------- 저장된 AI 계획 ----------
    def refresh_plans(self, select=None):
        import coach
        try:
            plans = [p for p in coach.list_plans() if str(p.get("id", "")).startswith("ai-")]
        except Exception as e:
            plans = []
            self.set_status(f"plan.json 을 못 읽었어요: {e}", RED)
        plans.sort(key=lambda p: p["date"], reverse=True)
        self._plans = {f"{p['date']}  {p.get('title', '')}  ({len(p.get('items', []))}개)": p["id"] for p in plans}
        labels = list(self._plans)
        self.plan_pick.configure(values=labels or ["(없음)"])
        pick = next((l for l, i in self._plans.items() if i == select), labels[0] if labels else "(없음)")
        self.plan_var.set(pick)

    def delete_selected(self):
        from tkinter import messagebox
        import coach
        label = self.plan_var.get()
        pid = self._plans.get(label)
        if not pid:
            return self.set_status("지울 AI 계획이 없어요", DIM)
        if not messagebox.askyesno("AI 계획 삭제", f"이 계획을 지울까요?\n\n{label}\n\n지우기 전 plan.json 은 백업 폴더에 남아요.", parent=self):
            return
        try:
            gone, bak = coach.delete_plan(pid)
        except Exception as e:
            return self.set_status(f"삭제 실패: {e}", RED)
        self.refresh_plans()
        tail = "폰 앱을 열면 카드도 사라져요" if self.httpd else "서버를 켜고 폰 앱을 열면 카드도 사라져요"
        self.set_status(f"✓ 삭제했어요 — {tail}" if gone else "이미 없는 계획이에요", GREEN if gone else DIM)
        print(f"AI 계획 삭제: {pid} (백업: {bak})")

    # ---------- 서버 ----------
    def start_server(self):
        import serve
        try:
            self.httpd = serve.make_server()
        except OSError:
            self.httpd = None
            self.dot.configure(fg=RED)
            self.srv_text.configure(text="서버를 못 켰어요 — 포트 8123 사용 중")
            self.url.configure(text="다른 서버 창(서버켜기.bat 등)이 켜져 있으면 닫고 [서버 켜기]를 누르세요")
            self.srv_btn.configure(text="서버 켜기")
            return
        threading.Thread(target=self.httpd.serve_forever, daemon=True).start()
        self.dot.configure(fg=GREEN)
        self.srv_text.configure(text="서버 켜짐")
        self.url.configure(text=f"폰에서 열기:  http://{serve.my_ip()}:{serve.PORT}   (같은 Wi-Fi)")
        self.srv_btn.configure(text="서버 끄기")
        print(f"서버 시작 — http://{serve.my_ip()}:{serve.PORT}")

    def stop_server(self):
        if self.httpd:
            self.httpd.shutdown()
            self.httpd.server_close()
            self.httpd = None
            print("서버 종료")
        self.dot.configure(fg=DIM)
        self.srv_text.configure(text="서버 꺼짐")
        self.url.configure(text="폰 동기화·계획 받기는 서버가 켜져 있어야 해요")
        self.srv_btn.configure(text="서버 켜기")

    def toggle_server(self):
        self.stop_server() if self.httpd else self.start_server()

    # ---------- 코치 ----------
    def run_coach(self, dry_run):
        if self.busy:
            return
        date = self.date.get().strip()
        try:
            datetime.date.fromisoformat(date)
        except ValueError:
            return self.set_status("날짜는 2026-10-08 형식으로", RED)
        try:
            minutes = int(self.time_var.get())
        except ValueError:
            return self.set_status("시간은 숫자(분)로", RED)
        focus = None if self.focus_var.get() == "자동" else self.focus_var.get()
        checked = [k for k, v in self.cond_vars.items() if v.get()]
        note = ", ".join(checked + [self.note.get("1.0", "end").strip()]).strip(", ")

        self.busy = True
        for btn in (self.make_btn, self.preview_btn):
            btn.state(["disabled"])
        self.set_status("계획 만드는 중… (보통 20~30초)", ORANGE)

        def work():
            import coach
            try:
                r = coach.make_plan(date, focus, note, minutes, dry_run)
                self.q.put(("done", r))
            except coach.CoachError as e:
                self.q.put(("fail", str(e)))
            except Exception as e:
                traceback.print_exc()
                self.q.put(("fail", f"오류: {e}"))
        threading.Thread(target=work, daemon=True).start()

    def finish(self, kind, payload):
        self.busy = False
        for btn in (self.make_btn, self.preview_btn):
            btn.state(["!disabled"])
        if kind == "fail":
            return self.set_status(payload, RED)
        self.show_result(format_plan(payload))
        n = sum(rule != "G6" for rule, _ in payload["fixes"])
        fix = f" · 자동수정 {n}건" if n else ""
        if payload["saved"]:
            self.refresh_plans(select=payload["plan"]["id"])
            tail ="폰 앱을 열면 홈에 카드가 떠요" if self.httpd else "서버를 켜고 폰 앱을 열면 카드가 떠요"
            self.set_status(f"✓ 저장했어요{fix} — {tail}", GREEN)
        else:
            self.set_status(f"미리보기{fix} — 마음에 들면 [계획 만들기]", DIM)

    # ---------- 공통 ----------
    def poll(self):
        try:
            while True:
                kind, payload = self.q.get_nowait()
                if kind == "log":
                    self.append_log(payload)
                else:
                    self.finish(kind, payload)
        except queue.Empty:
            pass
        self.after(100, self.poll)

    def report_callback_exception(self, exc, val, tb):
        """버튼 등 창 안에서 난 예외 — pythonw 라 콘솔이 없으니 파일과 상태줄로"""
        write_error("".join(traceback.format_exception(exc, val, tb)))
        self.set_status(f"오류: {val} (coach_log/panel_error.log)", RED)

    def on_close(self):
        self.stop_server()
        self.destroy()


if __name__ == "__main__":
    try:                                       # 고해상도 화면에서 글씨가 흐려지지 않게
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass
    try:
        Panel().mainloop()
    except Exception:                          # 시작하다 죽으면 pythonw 는 아무것도 안 보여 줌 → 파일 + 알림 창
        err = traceback.format_exc()
        write_error(err)
        try:
            from tkinter import messagebox
            r = tk.Tk(); r.withdraw()
            messagebox.showerror("헬스일지", f"시작하지 못했어요.\n\n{err[-600:]}\n\n기록: {ERROR_LOG}")
        except Exception:
            pass
