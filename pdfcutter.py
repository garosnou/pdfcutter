import os
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from pypdf import PdfReader, PdfWriter

try:
    from tkinterdnd2 import COPY, DND_FILES, TkinterDnD
except ImportError:
    TkinterDnD = None


# === Палитра ===
BG = "#f3f4f8"
CARD = "#ffffff"
TEXT = "#1f2937"
MUTED = "#6b7280"
BORDER = "#d7dbe3"
ACCENT = "#4f46e5"
ACCENT_ACTIVE = "#4338ca"
ACCENT_SOFT = "#eef0ff"
DISABLED = "#a5b4fc"
SUCCESS = "#15803d"
SUCCESS_SOFT = "#effaf2"
ERROR = "#dc2626"

FONT = "Segoe UI"


# === Логика разбиения ===
def parse_start_pages(text):
    """Возвращает (отсортированные уникальные номера ≥ 1, список нераспознанных фрагментов)."""
    pages, bad = set(), []
    for token in text.replace(";", ",").replace(" ", ",").split(","):
        token = token.strip()
        if not token:
            continue
        if token.isdigit() and int(token) >= 1:
            pages.add(int(token))
        else:
            bad.append(token)
    return sorted(pages), bad


def compute_ranges(total, start_pages, trim_covers):
    """Возвращает список (start, end, display_start) в абсолютной 1-based нумерации.

    При обрезке обложек исключаются страницы 1, 2 и две последние, а номера
    start_pages считаются относительными: 1 соответствует абсолютной 3-й странице.
    """
    if trim_covers:
        if total < 5:
            raise ValueError("Слишком мало страниц для обрезки обложек (нужно ≥ 5).")
        offset, work_end = 2, total - 2
    else:
        offset, work_end = 0, total

    starts = [p + offset for p in start_pages if p + offset <= work_end]
    ranges = []
    for i, start in enumerate(starts):
        end = starts[i + 1] - 1 if i + 1 < len(starts) else work_end
        ranges.append((start, end, start - offset))
    return ranges


def write_parts(pdf_path, output_dir, ranges, on_progress):
    reader = PdfReader(pdf_path)
    os.makedirs(output_dir, exist_ok=True)
    for i, (start, end, display_start) in enumerate(ranges, 1):
        writer = PdfWriter()
        for p in range(start - 1, end):
            writer.add_page(reader.pages[p])
        with open(os.path.join(output_dir, f"start_{display_start}.pdf"), "wb") as f:
            writer.write(f)
        on_progress(i)


def plural(n, one, few, many):
    if n % 10 == 1 and n % 100 != 11:
        return one
    if 2 <= n % 10 <= 4 and not 12 <= n % 100 <= 14:
        return few
    return many


# === GUI ===
class App:
    def __init__(self, root):
        self.root = root
        self.scale = root.winfo_fpixels("1i") / 96
        self.pdf_path = None
        self.total_pages = 0
        self.ranges = []
        self.busy = False
        self.drag_hover = False
        self.last_output_dir = None

        root.title("PDF Cutter")
        root.configure(bg=BG)
        root.minsize(self.px(560), self.px(700))
        root.geometry(f"{self.px(620)}x{self.px(700)}")

        self._init_styles()
        self._build()
        self._enable_dnd()
        self._update_preview()

    def px(self, value):
        return int(value * self.scale)

    # --- Оформление ---
    def _init_styles(self):
        style = ttk.Style(self.root)
        style.theme_use("clam")

        style.configure(".", background=BG, foreground=TEXT, font=(FONT, 10))
        style.configure("Card.TFrame", background=CARD)
        style.configure("Title.TLabel", background=BG, font=(FONT, 18, "bold"))
        style.configure("Subtitle.TLabel", background=BG, foreground=MUTED)
        style.configure("Field.TLabel", background=CARD, font=(FONT, 10, "bold"))
        style.configure("Hint.TLabel", background=CARD, foreground=MUTED, font=(FONT, 9))
        style.configure("Preview.TLabel", background=CARD, foreground=MUTED, font=(FONT, 9))
        style.configure("Status.TLabel", background=BG, foreground=MUTED)

        style.configure(
            "TEntry",
            fieldbackground=CARD,
            bordercolor=BORDER,
            lightcolor=BORDER,
            darkcolor=BORDER,
            padding=(self.px(8), self.px(6)),
        )
        style.map("TEntry", bordercolor=[("focus", ACCENT)], lightcolor=[("focus", ACCENT)])

        style.configure(
            "TButton",
            background=CARD,
            bordercolor=BORDER,
            lightcolor=CARD,
            darkcolor=CARD,
            focuscolor=CARD,
            padding=(self.px(12), self.px(6)),
        )
        style.map("TButton", background=[("active", ACCENT_SOFT)], bordercolor=[("active", ACCENT)])

        style.configure(
            "Accent.TButton",
            background=ACCENT,
            foreground="white",
            bordercolor=ACCENT,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            focuscolor=ACCENT,
            font=(FONT, 11, "bold"),
            padding=(self.px(24), self.px(10)),
        )
        style.map(
            "Accent.TButton",
            background=[("disabled", DISABLED), ("pressed", ACCENT_ACTIVE), ("active", ACCENT_ACTIVE)],
            bordercolor=[("disabled", DISABLED), ("active", ACCENT_ACTIVE)],
            lightcolor=[("disabled", DISABLED), ("active", ACCENT_ACTIVE)],
            darkcolor=[("disabled", DISABLED), ("active", ACCENT_ACTIVE)],
            foreground=[("disabled", "white")],
        )

        style.configure(
            "TCheckbutton",
            background=CARD,
            focuscolor=CARD,
            indicatorbackground=CARD,
            indicatorforeground="white",
            upperbordercolor=BORDER,
            lowerbordercolor=BORDER,
        )
        style.map(
            "TCheckbutton",
            background=[("active", CARD)],
            indicatorbackground=[("selected", ACCENT)],
            upperbordercolor=[("selected", ACCENT)],
            lowerbordercolor=[("selected", ACCENT)],
        )

        style.configure(
            "Accent.Horizontal.TProgressbar",
            background=ACCENT,
            troughcolor="#e5e7eb",
            bordercolor=BG,
            lightcolor=ACCENT,
            darkcolor=ACCENT,
            thickness=self.px(6),
        )

    def _build(self):
        pad = self.px(20)
        outer = ttk.Frame(self.root, padding=(pad, self.px(16), pad, pad))
        outer.pack(fill="both", expand=True)

        ttk.Label(outer, text="PDF Cutter", style="Title.TLabel").pack(anchor="w")
        ttk.Label(outer, text="Разбивает PDF на части по начальным страницам", style="Subtitle.TLabel").pack(
            anchor="w", pady=(0, self.px(12))
        )

        self.drop = tk.Canvas(outer, height=self.px(150), bg=BG, highlightthickness=0, cursor="hand2")
        self.drop.pack(fill="x")
        self.drop.bind("<Configure>", lambda e: self._draw_drop_zone())
        self.drop.bind("<Button-1>", lambda e: self.browse_pdf())

        card = ttk.Frame(outer, style="Card.TFrame", padding=self.px(16))
        card.pack(fill="x", pady=(self.px(14), 0))
        card.columnconfigure(0, weight=1)

        ttk.Label(card, text="Папка вывода", style="Field.TLabel").grid(row=0, column=0, columnspan=2, sticky="w")
        self.var_output = tk.StringVar()
        ttk.Entry(card, textvariable=self.var_output).grid(row=1, column=0, sticky="ew", pady=(self.px(4), 0))
        ttk.Button(card, text="Обзор…", command=self.browse_output).grid(
            row=1, column=1, padx=(self.px(8), 0), pady=(self.px(4), 0)
        )

        ttk.Label(card, text="Начальные страницы", style="Field.TLabel").grid(
            row=2, column=0, columnspan=2, sticky="w", pady=(self.px(14), 0)
        )
        self.var_pages = tk.StringVar(value="1,5,12,20")
        self.var_pages.trace_add("write", lambda *_: self._update_preview())
        ttk.Entry(card, textvariable=self.var_pages).grid(
            row=3, column=0, columnspan=2, sticky="ew", pady=(self.px(4), 0)
        )
        ttk.Label(card, text="Через запятую, например: 1, 5, 12, 20", style="Hint.TLabel").grid(
            row=4, column=0, columnspan=2, sticky="w", pady=(self.px(2), 0)
        )

        self.var_trim = tk.BooleanVar(value=False)
        self.var_trim.trace_add("write", lambda *_: self._update_preview())
        ttk.Checkbutton(
            card, text="Обрезать обложки (1, 2 и два последних листа)", variable=self.var_trim
        ).grid(row=5, column=0, columnspan=2, sticky="w", pady=(self.px(12), 0))

        self.preview = ttk.Label(card, style="Preview.TLabel", justify="left")
        self.preview.grid(row=6, column=0, columnspan=2, sticky="w", pady=(self.px(12), 0))
        card.bind("<Configure>", lambda e: self.preview.configure(wraplength=e.width - self.px(32)))

        actions = ttk.Frame(outer)
        actions.pack(fill="x", pady=(self.px(16), 0))
        self.btn_split = ttk.Button(actions, text="Нарезать PDF", style="Accent.TButton", command=self.split)
        self.btn_split.pack(side="left")
        self.btn_open = ttk.Button(actions, text="Открыть папку", command=self.open_output)

        self.progress = ttk.Progressbar(outer, style="Accent.Horizontal.TProgressbar", mode="determinate")
        self.progress.pack(fill="x", pady=(self.px(14), 0))
        self.status = ttk.Label(outer, style="Status.TLabel")
        self.status.pack(anchor="w", pady=(self.px(6), 0))

        self.root.bind("<Return>", lambda e: self.split())

    # --- Зона перетаскивания ---
    def _rounded_rect(self, x1, y1, x2, y2, r, **kw):
        points = [
            x1 + r, y1, x2 - r, y1, x2, y1, x2, y1 + r,
            x2, y2 - r, x2, y2, x2 - r, y2, x1 + r, y2,
            x1, y2, x1, y2 - r, x1, y1 + r, x1, y1,
        ]
        return self.drop.create_polygon(points, smooth=True, **kw)

    def _draw_drop_zone(self):
        c = self.drop
        c.delete("all")
        w, h = c.winfo_width(), c.winfo_height()
        if w < 50:
            return

        if self.drag_hover:
            fill, outline, accent = ACCENT_SOFT, ACCENT, ACCENT
            title, subtitle = "Отпустите, чтобы открыть файл", ""
        elif self.pdf_path:
            fill, outline, accent = SUCCESS_SOFT, SUCCESS, SUCCESS
            title = os.path.basename(self.pdf_path)
            subtitle = (
                f"{self.total_pages} {plural(self.total_pages, 'страница', 'страницы', 'страниц')}"
                " · перетащите другой файл или нажмите, чтобы выбрать"
            )
        else:
            fill, outline, accent = CARD, BORDER, MUTED
            title = "Перетащите PDF сюда" if TkinterDnD else "Нажмите, чтобы выбрать PDF"
            subtitle = "или нажмите, чтобы выбрать файл" if TkinterDnD else ""

        m = self.px(2)
        self._rounded_rect(m, m, w - m, h - m, self.px(16), fill=fill, outline=outline, width=self.px(2))

        # Иконка документа
        cx, top = w // 2, self.px(24)
        dw, dh, fold = self.px(34), self.px(44), self.px(11)
        x1, x2 = cx - dw // 2, cx + dw // 2
        c.create_polygon(
            x1, top, x2 - fold, top, x2, top + fold, x2, top + dh, x1, top + dh,
            fill=CARD, outline=accent, width=self.px(2), joinstyle="round",
        )
        c.create_line(x2 - fold, top, x2 - fold, top + fold, x2, top + fold, fill=accent, width=self.px(2))
        c.create_text(cx, top + dh * 0.65, text="PDF", fill=accent, font=(FONT, 8, "bold"))

        c.create_text(cx, top + dh + self.px(22), text=title, fill=TEXT, font=(FONT, 11, "bold"), width=w - self.px(40))
        if subtitle:
            c.create_text(cx, top + dh + self.px(46), text=subtitle, fill=MUTED, font=(FONT, 9), width=w - self.px(40))

    def _enable_dnd(self):
        if not TkinterDnD:
            return
        self.root.drop_target_register(DND_FILES)
        self.root.dnd_bind("<<DropEnter>>", self._on_drag_enter)
        self.root.dnd_bind("<<DropPosition>>", lambda e: COPY)
        self.root.dnd_bind("<<DropLeave>>", self._on_drag_leave)
        self.root.dnd_bind("<<Drop>>", self._on_drop)

    def _on_drag_enter(self, event):
        self.drag_hover = True
        self._draw_drop_zone()
        return COPY

    def _on_drag_leave(self, event):
        self.drag_hover = False
        self._draw_drop_zone()

    def _on_drop(self, event):
        self.drag_hover = False
        files = self.root.tk.splitlist(event.data)
        pdf = next((f for f in files if f.lower().endswith(".pdf")), None)
        if pdf:
            self.load_pdf(pdf)
        else:
            self._draw_drop_zone()
            self.set_status("Это не PDF-файл", ERROR)
        return COPY

    # --- Действия ---
    def load_pdf(self, path):
        if self.busy:
            return
        path = os.path.normpath(path)
        try:
            total = len(PdfReader(path).pages)
        except Exception as e:
            self.set_status(f"Не удалось открыть PDF: {e}", ERROR)
            self._draw_drop_zone()
            return
        self.pdf_path, self.total_pages = path, total
        self.var_output.set(os.path.dirname(path))
        self.progress["value"] = 0
        self.btn_open.pack_forget()
        self.set_status("")
        self._draw_drop_zone()
        self._update_preview()

    def browse_pdf(self):
        path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
        if path:
            self.load_pdf(path)

    def browse_output(self):
        path = filedialog.askdirectory(initialdir=self.var_output.get() or None)
        if path:
            self.var_output.set(os.path.normpath(path))

    def open_output(self):
        if self.last_output_dir and os.path.isdir(self.last_output_dir):
            os.startfile(self.last_output_dir)

    def _update_preview(self):
        self.ranges = []
        if not self.pdf_path:
            self.preview.configure(text="Выберите PDF-файл, чтобы увидеть, какие части получатся.", foreground=MUTED)
            self._refresh_button()
            return

        pages, bad = parse_start_pages(self.var_pages.get())
        try:
            self.ranges = compute_ranges(self.total_pages, pages, self.var_trim.get())
        except ValueError as e:
            self.preview.configure(text=str(e), foreground=ERROR)
            self._refresh_button()
            return

        notes = []
        if bad:
            notes.append("Не распознано: " + ", ".join(bad))
        skipped = len(pages) - len(self.ranges)
        if skipped:
            limit = self.total_pages - 4 if self.var_trim.get() else self.total_pages
            notes.append(f"Пропущено начальных страниц за пределами документа: {skipped} (доступно 1–{limit})")

        if self.ranges:
            n = len(self.ranges)
            parts = []
            for start, end, shown in self.ranges:
                last = shown + end - start
                parts.append(f"{shown}–{last}" if last != shown else str(shown))
            text = f"Будет создано {n} {plural(n, 'файл', 'файла', 'файлов')}: стр. {', '.join(parts)}"
            self.preview.configure(text="\n".join([text] + notes), foreground=ERROR if notes else MUTED)
        else:
            self.preview.configure(text="\n".join(notes or ["Укажите хотя бы одну начальную страницу."]), foreground=ERROR)
        self._refresh_button()

    def _refresh_button(self):
        ready = bool(self.ranges) and not self.busy
        self.btn_split.state(["!disabled"] if ready else ["disabled"])

    def set_status(self, text, color=MUTED):
        self.status.configure(text=text, foreground=color)

    def split(self):
        if self.busy or not self.ranges:
            return
        output_dir = self.var_output.get().strip() or os.path.dirname(self.pdf_path)
        self.var_output.set(output_dir)

        self.busy = True
        self._refresh_button()
        self.btn_open.pack_forget()
        self.progress.configure(maximum=len(self.ranges), value=0)
        self.set_status("Нарезаю…")

        ranges = list(self.ranges)

        def on_progress(i):
            self.root.after(0, lambda: self._on_progress(i, len(ranges)))

        def worker():
            try:
                write_parts(self.pdf_path, output_dir, ranges, on_progress)
            except Exception as e:
                self.root.after(0, lambda: self._on_finished(output_dir, len(ranges), e))
            else:
                self.root.after(0, lambda: self._on_finished(output_dir, len(ranges), None))

        threading.Thread(target=worker, daemon=True).start()

    def _on_progress(self, i, n):
        self.progress["value"] = i
        self.set_status(f"Нарезаю… {i} из {n}")

    def _on_finished(self, output_dir, n, error):
        self.busy = False
        self._refresh_button()
        if error:
            self.progress["value"] = 0
            self.set_status("Ошибка при нарезке", ERROR)
            messagebox.showerror("Ошибка", f"Не удалось обработать PDF:\n{error}")
            return
        self.last_output_dir = output_dir
        self.set_status(f"Готово: {n} {plural(n, 'файл', 'файла', 'файлов')} в папке {output_dir}", SUCCESS)
        self.btn_open.pack(side="left", padx=(self.px(10), 0))


def main():
    if sys.platform == "win32":
        try:
            import ctypes
            ctypes.windll.shcore.SetProcessDpiAwareness(1)
        except Exception:
            pass

    global TkinterDnD
    root = None
    if TkinterDnD:
        try:
            root = TkinterDnD.Tk()
        except Exception:
            # Tk уже создан до ошибки загрузки tkdnd — закрываем его, чтобы не висело пустое окно
            if tk._default_root:
                tk._default_root.destroy()
            TkinterDnD = None
    if root is None:
        root = tk.Tk()
    app = App(root)

    # PDF, брошенный на иконку .exe, приходит аргументом командной строки
    if len(sys.argv) > 1 and sys.argv[1].lower().endswith(".pdf") and os.path.isfile(sys.argv[1]):
        root.after(100, lambda: app.load_pdf(sys.argv[1]))

    root.mainloop()


if __name__ == "__main__":
    main()
