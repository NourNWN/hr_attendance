import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import os
import webbrowser

from hr_attendance import phase1, phase2

# ═════════════════════════════════════════════════════════════
#  THEME & STYLES
# ═════════════════════════════════════════════════════════════
BG         = "#F0F4F3"
CARD       = "#FFFFFF"
TEAL       = "#00695C"
TEAL_DARK  = "#004D40"
TEAL_LIGHT = "#E0F2F1"
ACCENT     = "#F57F17"
TEXT       = "#1C1C1C"
MUTED      = "#6B7280"
BORDER     = "#D1D5DB"
INPUT_BG   = "#F9FAFB"

F_TITLE = ("Georgia", 18, "bold")
F_SUB   = ("Helvetica", 9)
F_STEP  = ("Helvetica", 8, "bold")
F_HEAD  = ("Helvetica", 11, "bold")
F_LABEL = ("Helvetica", 9)
F_ENTRY = ("Helvetica", 10)
F_BTN   = ("Helvetica", 10, "bold")

# ═════════════════════════════════════════════════════════════
#  HELPERS
# ═════════════════════════════════════════════════════════════

def make_entry(parent, var, w=10):
    # تحسين: إضافة justify="center" لتوسيط القيم
    e = tk.Entry(parent, textvariable=var, font=F_ENTRY,
                 fg=TEXT, bg=INPUT_BG, relief="flat",
                 highlightbackground=BORDER, highlightcolor=TEAL,
                 highlightthickness=1, insertbackground=TEAL,
                 width=w, justify="center")
    # تحسين: جعل المربعات تتمدد لملء الفراغ
    e.pack(side="left", ipady=5, fill="x", expand=True)
    return e

def make_btn(parent, text, cmd, color=TEAL):
    b = tk.Button(parent, text=text, command=cmd, font=F_BTN,
                  fg="white", bg=color, activebackground=TEAL_DARK,
                  activeforeground="white", relief="flat",
                  cursor="hand2", pady=9, bd=0)
    b.bind("<Enter>", lambda e: b.config(bg=TEAL_DARK))
    b.bind("<Leave>", lambda e: b.config(bg=color))
    return b

def sep(parent, pady=10):
    tk.Frame(parent, height=1, bg=BORDER).pack(fill="x", pady=pady)

def row_label(parent, text):
    tk.Label(parent, text=text, font=F_LABEL, fg=MUTED,
             bg=CARD, anchor="w").pack(anchor="w", pady=(6, 2))

# ═════════════════════════════════════════════════════════════
#  ToolTip CLASS
# ═════════════════════════════════════════════════════════════
class ToolTip:
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        self.widget.bind("<Enter>", self.show_tip)
        self.widget.bind("<Leave>", self.hide_tip)

    def show_tip(self, event=None):
        x, y, cx, cy = self.widget.bbox("insert")
        x += self.widget.winfo_rootx() + 25
        y += self.widget.winfo_rooty() + 20
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(1)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify='left',
                         background="#ffffe0", relief='solid', borderwidth=1,
                         font=("tahoma", "8", "normal"))
        label.pack(ipadx=1)

    def hide_tip(self, event=None):
        if self.tipwindow: self.tipwindow.destroy()
        self.tipwindow = None

# ═════════════════════════════════════════════════════════════
#  APP CLASS
# ═════════════════════════════════════════════════════════════

class HRApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("HR Attendance Calculator")
        self.configure(bg=BG)
        self.resizable(False, False)
        self._center(820, 620)

        # تحسين: استخدام أيقونة النظام المدمجة (شكل ترس الإعدادات)
        try:
            self.iconbitmap("shell32.dll")
        except:
            pass

        self.p1_file    = tk.StringVar()
        self.p2_file    = tk.StringVar()
        self.work_start = tk.StringVar(value="09:00")
        self.work_end   = tk.StringVar(value="17:00")
        self.grace      = tk.StringVar(value="8")
        self.allowance  = tk.StringVar(value="15")

        self._build()

    def _center(self, w, h):
        self.update_idletasks()
        x = (self.winfo_screenwidth()  - w) // 2
        y = (self.winfo_screenheight() - h) // 2
        self.geometry(f"{w}x{h}+{x}+{y}")

    def _validate_num(self, P):
        return P == "" or all(c in "0123456789:" for c in P)

    def _add_tooltip(self, widget, text):
        def show(event):
            self.tip = tk.Toplevel()
            self.tip.wm_overrideredirect(True)  # حذف حواف النافذة
            self.tip.geometry(f"+{event.x_root + 10}+{event.y_root + 10}")
            tk.Label(self.tip, text=text, bg="#FFFFE0", relief="solid",
                     borderwidth=1, font=("Helvetica", 8)).pack()

        def hide(event):
            if hasattr(self, 'tip'): self.tip.destroy()

        widget.bind("<Enter>", show)
        widget.bind("<Leave>", hide)

    def _build(self):
        self._header()
        self._body()
        self._status_bar()

    def _header(self):
        h = tk.Frame(self, bg=TEAL, padx=24, pady=16)
        h.pack(fill="x")
        tk.Label(h, text="HR Attendance Calculator",
                 font=F_TITLE, fg="white", bg=TEAL).pack(anchor="w")
        tk.Label(h, text="Generate monthly attendance reports from fingerprint data",
                 font=F_SUB, fg=TEAL_LIGHT, bg=TEAL).pack(anchor="w", pady=(2, 0))

    def _body(self):
        outer = tk.Frame(self, bg=BG, padx=20, pady=16)
        outer.pack(fill="both", expand=True)

        cols = tk.Frame(outer, bg=BG)
        cols.pack(fill="both", expand=True)

        # تحسين: توحيد عرض الكاردات 50/50 باستخدام uniform
        cols.grid_columnconfigure(0, weight=1, uniform="group1")
        cols.grid_columnconfigure(1, minsize=16)
        cols.grid_columnconfigure(2, weight=1, uniform="group1")
        cols.grid_rowconfigure(0, weight=1)

        self._phase1_card(cols)
        self._phase2_card(cols)

    def _step_title(self, parent, step, title):
        row = tk.Frame(parent, bg=BG)
        row.pack(fill="x", pady=(0, 4))
        tk.Label(row, text=f" {step} ", font=F_STEP,
                 fg="white", bg=TEAL, padx=4, pady=3).pack(side="left")
        tk.Label(row, text=f"  {title}", font=F_HEAD,
                 fg=TEXT, bg=BG).pack(side="left")

    def _phase1_card(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.grid(row=0, column=0, sticky="nsew")
        self._step_title(wrap, "STEP 1", "Generate Review File")

        c = tk.Frame(wrap, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        c.pack(fill="both", expand=True)
        inner = tk.Frame(c, bg=CARD, padx=18, pady=16)
        inner.pack(fill="both", expand=True)

        # 1. تحسين استجابة Browse (جعل الحقل قابلاً للضغط)
        row_label(inner, "Fingerprint File (.xlsx)")
        fr = tk.Frame(inner, bg=CARD)
        fr.pack(fill="x", pady=(0, 4))

        ent_file = tk.Entry(fr, textvariable=self.p1_file, font=F_LABEL,
                            fg=TEXT, bg=INPUT_BG, relief="flat",
                            highlightbackground=BORDER, highlightthickness=1,
                            state="readonly", cursor="hand2")  # إضافة شكل اليد
        ent_file.pack(side="left", ipady=5, padx=(0, 8), fill="x", expand=True)
        # ربط الحقل بفتح نافذة الاختيار عند الضغط عليه
        ent_file.bind("<Button-1>", lambda e: self._browse1())

        tk.Button(fr, text="Browse", font=F_LABEL, fg=TEAL, bg=CARD,
                  relief="flat", cursor="hand2", bd=0,
                  command=self._browse1).pack(side="left")

        sep(inner)
        tk.Label(inner, text="Settings", font=F_HEAD, fg=TEXT, bg=CARD).pack(anchor="w", pady=(0, 8))

        # 2. إعداد الحماية (Validation) والتلميحات (Tooltips)
        vcmd = (self.register(self._validate_num), '%P')
        tips = {
            "Work Start Time": "Enter time in 24h format (e.g. 09:00)",
            "Work End Time": "Enter time in 24h format (e.g. 17:00)",
            "Grace Period (min)": "Extra minutes allowed before being marked late",
            "Monthly Late Allowance (min)": "Total free late minutes per month"
        }

        for lbl, var in [
            ("Work Start Time", self.work_start),
            ("Work End Time", self.work_end),
            ("Grace Period (min)", self.grace),
            ("Monthly Late Allowance (min)", self.allowance),
        ]:
            r = tk.Frame(inner, bg=CARD)
            r.pack(fill="x", pady=3)

            lbl_widget = tk.Label(r, text=lbl, font=F_LABEL, fg=MUTED,
                                  bg=CARD, width=25, anchor="w")
            lbl_widget.pack(side="left")

            # إضافة تلميح للأداة عند الوقوف على النص
            self._add_tooltip(lbl_widget, tips.get(lbl, ""))

            # إنشاء الحقل مع الحماية وتوسيط النص
            e = tk.Entry(r, textvariable=var, font=F_ENTRY, fg=TEXT, bg=INPUT_BG,
                         relief="flat", highlightbackground=BORDER, highlightthickness=1,
                         justify="center", validate="key", validatecommand=vcmd)
            e.pack(side="left", ipady=5, fill="x", expand=True)
            # إضافة تلميح للأداة عند الوقوف على المربع نفسه أيضاً
            self._add_tooltip(e, tips.get(lbl, ""))

        # 3. محاذاة أزرار Run في الأسفل
        tk.Frame(inner, bg=CARD).pack(fill="both", expand=True)
        sep(inner, pady=12)
        make_btn(inner, "▶    Run Phase 1", self._run1).pack(fill="x", side="bottom")

    def _phase2_card(self, parent):
        wrap = tk.Frame(parent, bg=BG)
        wrap.grid(row=0, column=2, sticky="nsew")
        self._step_title(wrap, "STEP 2", "Generate Final Report")

        c = tk.Frame(wrap, bg=CARD, highlightbackground=BORDER, highlightthickness=1)
        c.pack(fill="both", expand=True)
        inner = tk.Frame(c, bg=CARD, padx=18, pady=16)
        inner.pack(fill="both", expand=True)

        row_label(inner, "Review File (.xlsx)")
        fr = tk.Frame(inner, bg=CARD)
        fr.pack(fill="x", pady=(0, 4))
        tk.Entry(fr, textvariable=self.p2_file, font=F_LABEL,
                 fg=TEXT, bg=INPUT_BG, relief="flat",
                 highlightbackground=BORDER, highlightthickness=1,
                 state="readonly").pack(side="left", ipady=5, padx=(0, 8), fill="x", expand=True)
        tk.Button(fr, text="Browse", font=F_LABEL, fg=TEAL, bg=CARD,
                  relief="flat", cursor="hand2", bd=0,
                  command=self._browse2).pack(side="left")

        sep(inner)
        tk.Label(inner, text="Leave Type Reference", font=F_HEAD, fg=TEXT, bg=CARD).pack(anchor="w", pady=(0, 8))

        for code, color, title, desc in [
            ("AM", ACCENT, "Morning Leave",   "Check In time is ignored for that day"),
            ("PM", TEAL,   "Afternoon Leave", "Check Out time is ignored for that day"),
        ]:
            row = tk.Frame(inner, bg=INPUT_BG, highlightbackground=BORDER, highlightthickness=1)
            row.pack(fill="x", pady=4, ipady=8, ipadx=10)
            tk.Label(row, text=f"  {code}  ", font=("Courier", 11, "bold"), fg=color, bg=INPUT_BG).pack(side="left")
            sub = tk.Frame(row, bg=INPUT_BG); sub.pack(side="left")
            tk.Label(sub, text=title, font=("Helvetica", 9, "bold"), fg=TEXT, bg=INPUT_BG, anchor="w").pack(anchor="w")
            tk.Label(sub, text=desc, font=F_LABEL, fg=MUTED, bg=INPUT_BG, anchor="w").pack(anchor="w")

        note = tk.Frame(inner, bg=TEAL_LIGHT, highlightbackground="#B2DFDB", highlightthickness=1)
        note.pack(fill="x", pady=(12, 0), ipady=8, ipadx=10)
        tk.Label(note, text="ℹ Settings are loaded automatically from review file.",
                 font=F_LABEL, fg=TEAL_DARK, bg=TEAL_LIGHT, justify="left").pack(anchor="w", padx=8)

        # تحسين: Spacer لضمان محاذاة أزرار Run في الأسفل
        tk.Frame(inner, bg=CARD).pack(fill="both", expand=True)
        sep(inner, pady=12)
        make_btn(inner, "▶    Run Phase 2", self._run2, color=TEAL_DARK).pack(fill="x", side="bottom")

    def _status_bar(self):
        bar = tk.Frame(self, bg=TEAL_DARK, padx=16, pady=10)
        bar.pack(fill="x", side="bottom")
        self._s_icon = tk.Label(bar, text="●", font=("Helvetica", 11), fg=TEAL_LIGHT, bg=TEAL_DARK)
        self._s_icon.pack(side="left")
        self._s_msg = tk.Label(bar, text="  Ready", font=("Helvetica", 9), fg=TEAL_LIGHT, bg=TEAL_DARK, anchor="w")
        self._s_msg.pack(side="left", fill="x", expand=True)

        self.dev_label = tk.Label(bar, text="Designed & Engineered by Eng. Nour Nasser",
                                  font=("Helvetica", 8, "italic"), fg=TEAL_LIGHT, bg=TEAL_DARK, cursor="hand2")
        self.dev_label.pack(side="right")
        self.dev_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.linkedin.com/in/nour-nasser-nwn"))
        self.dev_label.bind("<Enter>", lambda e: self.dev_label.config(fg="white"))
        self.dev_label.bind("<Leave>", lambda e: self.dev_label.config(fg=TEAL_LIGHT))

    def _set_status(self, msg, kind="info"):
        colors = {"info": TEAL_LIGHT, "ok": "#A5D6A7", "err": "#EF9A9A"}
        icons  = {"info": "●", "ok": "✔", "err": "✘"}
        self._s_icon.config(text=icons.get(kind, "●"), fg=colors.get(kind, TEAL_LIGHT))
        self._s_msg.config(text=f"  {msg}", fg=colors.get(kind, TEAL_LIGHT))

    def _browse1(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if p: self.p1_file.set(p)

    def _browse2(self):
        p = filedialog.askopenfilename(filetypes=[("Excel files", "*.xlsx *.xls")])
        if p: self.p2_file.set(p)

    def _run1(self):
        if not self.p1_file.get():
            messagebox.showwarning("Missing File", "Please select a fingerprint file.")
            return
        threading.Thread(target=self._exec1, daemon=True).start()

    def _exec1(self):
        self._set_status("Running Phase 1…", "info")
        try:
            out = phase1(self.p1_file.get(), self.work_start.get().strip(), self.work_end.get().strip(),
                         int(self.grace.get().strip()), int(self.allowance.get().strip()))
            if out:
                self.p2_file.set(out)
                # تحسين: رسالة ثم فتح الملف تلقائياً
                messagebox.showinfo("Review Required", f"Phase 1 complete!\n\nReview file '{os.path.basename(out)}' will open now.\nPlease fill Leave Types, SAVE, and then Run Phase 2.")
                os.startfile(out) if hasattr(os, 'startfile') else os.system(f'open "{out}"')
                self._set_status(f"Phase 1 complete - Review file opened.", "ok")
        except Exception as ex: self._set_status(f"Error: {ex}", "err")

    def _run2(self):
        if not self.p2_file.get():
            messagebox.showwarning("Missing File", "Please select a review file.")
            return
        threading.Thread(target=self._exec2, daemon=True).start()

    def _exec2(self):
        self._set_status("Running Phase 2…", "info")
        try:
            phase2(self.p2_file.get())
            full_path = os.path.abspath(os.path.dirname(self.p2_file.get()) or ".")
            self._set_status(f"Phase 2 complete — Saved in: {full_path}", "ok")
        except Exception as ex: self._set_status(f"Error: {ex}", "err")
