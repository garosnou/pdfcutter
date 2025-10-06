import os
import sys
import tkinter as tk
from tkinter import filedialog, messagebox
from pypdf import PdfReader, PdfWriter
import threading


def split_pdf():
    pdf_path = entry_pdf.get()
    output_dir = entry_output.get()
    pages_str = entry_pages.get().strip()
    trim = var_trim.get()

    if not pdf_path or not os.path.isfile(pdf_path):
        messagebox.showerror("Ошибка", "Укажите корректный PDF-файл.")
        return

    if not output_dir:
        output_dir = os.path.dirname(pdf_path)
        entry_output.delete(0, tk.END)
        entry_output.insert(0, output_dir)

    if not os.path.isdir(output_dir):
        try:
            os.makedirs(output_dir, exist_ok=True)
        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось создать папку вывода:\n{e}")
            return

    if not pages_str:
        messagebox.showerror("Ошибка", "Введите начальные страницы (через запятую).")
        return

    # Запускаем в фоне, чтобы не зависало окно
    threading.Thread(target=do_split, args=(pdf_path, output_dir, pages_str, trim), daemon=True).start()


def do_split(pdf_path, output_dir, pages_str, trim_covers=False):
    try:
        reader = PdfReader(pdf_path)
        total = len(reader.pages)
        
        # Рассчитываем рабочий диапазон для разбиения
        # По умолчанию: весь документ 1..total (1-based)
        work_start_abs = 1
        work_end_abs = total
        if trim_covers:
            # Обрезаем 1,2 и два последних листа
            if total < 5:
                raise ValueError("Слишком мало страниц для обрезки обложек (нужно ≥ 5).")
            work_start_abs = 3              # начинаем с 3-й страницы
            work_end_abs = total - 2        # заканчиваем на предпоследней-1 (исключая 2 последних)

        # Парсим начальные страницы
        start_pages_input = sorted(set(int(x.strip()) for x in pages_str.split(',') if x.strip().isdigit()))
        if not start_pages_input or min(start_pages_input) < 1:
            raise ValueError("Номера страниц должны быть ≥ 1.")

        # Преобразуем относительные номера в абсолютные, если включена обрезка
        if trim_covers:
            inner_count = work_end_abs - work_start_abs + 1
            if max(start_pages_input) > inner_count:
                messagebox.showwarning(
                    "Предупреждение",
                    f"Некоторые страницы выходят за пределы внутреннего диапазона (всего {inner_count} стр. после обрезки)."
                )
            start_pages_abs = []
            for rel in start_pages_input:
                if rel <= inner_count:
                    # Абсолютный 1-based старт = (0-based 2) + rel
                    start_pages_abs.append(2 + rel)
            start_pages = start_pages_abs
        else:
            if max(start_pages_input) > total:
                messagebox.showwarning("Предупреждение", f"Некоторые страницы выходят за пределы документа (всего {total} стр.).")
            start_pages = start_pages_input

        if not start_pages:
            raise ValueError("Нет валидных начальных страниц для разбиения.")

        # Формируем диапазоны внутри рабочего интервала
        ranges = []
        for i, start in enumerate(start_pages):
            end = start_pages[i + 1] - 1 if i + 1 < len(start_pages) else work_end_abs
            # Гарантируем, что не выходим за рабочие границы
            start = max(start, work_start_abs)
            end = min(end, work_end_abs)
            if start <= end:
                ranges.append((start, end))

        for start, end in ranges:
            if start > total:
                continue
            writer = PdfWriter()
            for p in range(start - 1, min(end, total)):
                writer.add_page(reader.pages[p])
            # При обрезке обложек сохраняем относительную нумерацию (1 соответствует абсолютной 3)
            display_start = start - (work_start_abs - 1) if trim_covers else start
            output_path = os.path.join(output_dir, f"start_{display_start}.pdf")
            with open(output_path, "wb") as f:
                writer.write(f)

        messagebox.showinfo("Готово", f"Создано {len(ranges)} файлов в папке:\n{output_dir}")

    except Exception as e:
        messagebox.showerror("Ошибка", f"Не удалось обработать PDF:\n{e}")


def browse_pdf():
    path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if path:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, path)
        # Автоматически ставим папку вывода = папка файла
        entry_output.delete(0, tk.END)
        entry_output.insert(0, os.path.dirname(path))


def browse_output():
    path = filedialog.askdirectory()
    if path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, path)


# === Поддержка перетаскивания файла (Windows) ===
def enable_drag_and_drop(window):
    try:
        import windnd
        def on_drop(files):
            for file in files:
                if isinstance(file, bytes):
                    try:
                        # Надёжно декодируем путь, учитывая системную кодировку файловой системы Windows
                        file = os.fsdecode(file)
                    except Exception:
                        import locale
                        enc = locale.getpreferredencoding(False) or 'mbcs'
                        file = file.decode(enc, errors='surrogateescape')
                if file.lower().endswith('.pdf'):
                    entry_pdf.delete(0, tk.END)
                    entry_pdf.insert(0, file)
                    entry_output.delete(0, tk.END)
                    entry_output.insert(0, os.path.dirname(file))
                    break
        windnd.hook_dropfiles(window.winfo_id(), func=on_drop)
    except ImportError:
        pass  # Если windnd не установлен — просто игнорируем


# === GUI ===
root = tk.Tk()
root.title("PDF Splitter — перетащи PDF сюда!")
root.geometry("650x260")
root.resizable(False, False)

tk.Label(root, text="PDF-файл:").grid(row=0, column=0, padx=10, pady=10, sticky="w")
entry_pdf = tk.Entry(root, width=65)
entry_pdf.grid(row=0, column=1, padx=5, pady=10)
tk.Button(root, text="Выбрать", command=browse_pdf).grid(row=0, column=2, padx=5)

tk.Label(root, text="Папка вывода:").grid(row=1, column=0, padx=10, pady=10, sticky="w")
entry_output = tk.Entry(root, width=65)
entry_output.grid(row=1, column=1, padx=5, pady=10)
tk.Button(root, text="Выбрать", command=browse_output).grid(row=1, column=2, padx=5)

tk.Label(root, text="Начальные страницы (через запятую):").grid(row=2, column=0, padx=10, pady=10, sticky="w")
entry_pages = tk.Entry(root, width=65)
entry_pages.grid(row=2, column=1, columnspan=2, padx=5, pady=10)
entry_pages.insert(0, "1,5,12,20")

# Чекбокс "Обрезать обложки": исключить 1,2 и два последних листа из разбиения
var_trim = tk.BooleanVar(value=False)
chk_trim = tk.Checkbutton(root, text="Обрезать обложки (1,2 и два последних листа)", variable=var_trim)
chk_trim.grid(row=3, column=1, columnspan=2, sticky="w", padx=5)

tk.Button(root, text="Нарезать PDF", command=split_pdf, bg="#4CAF50", fg="white", height=2, font=("Arial", 10, "bold")).grid(
    row=4, column=1, pady=20
)

# Включаем перетаскивание (только на Windows)
if sys.platform == "win32":
    try:
        import windnd
        enable_drag_and_drop(root)
    except ImportError:
        pass

root.mainloop()