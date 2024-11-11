import os
import sqlite3
from datetime import datetime
import getpass
from tkinter import Tk, filedialog, Label, Button, Entry, messagebox
from tkinter import ttk
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from rapidfuzz import fuzz  # Импортируем библиотеку для анализа схожести строк

# Функция для подключения к базе данных
def connect_to_database(db_path):
    global conn, cursor
    if os.path.exists(db_path):
        os.remove(db_path)  # Удаляем существующую базу, чтобы создать новую с правильной структурой

    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS files (
                        id INTEGER PRIMARY KEY,
                        parent_folder TEXT,
                        path TEXT,
                        filename TEXT,
                        last_modified TEXT,
                        created_by TEXT
                    )''')
    conn.commit()
    messagebox.showinfo("Info", "Database connected successfully!")

# Функция для добавления или обновления файла в базе
def update_file_in_db(parent_folder, path, filename, last_modified, created_by):
    cursor.execute('''INSERT OR REPLACE INTO files (parent_folder, path, filename, last_modified, created_by)
                      VALUES (?, ?, ?, ?, ?)''', (parent_folder, path, filename, last_modified, created_by))
    conn.commit()
    update_table()

# Функция для проверки схожести названий файлов
def are_filenames_similar(name1, name2, threshold=80):
    # Убираем расширение и сравниваем только базовые имена
    base1, _ = os.path.splitext(name1)
    base2, _ = os.path.splitext(name2)
    similarity = fuzz.ratio(base1, base2)
    return similarity >= threshold

# Функция для подсветки файлов с похожими именами и одинаковой датой
def apply_highlighting():
    rows = cursor.execute("SELECT * FROM files").fetchall()
    file_groups = {}

    # Группируем файлы по папке
    for row in rows:
        parent_folder, path, filename, last_modified, created_by = row[1:]
        key = (parent_folder, last_modified[:10])  # Группируем по папке и дате изменения

        # Группируем файлы по папке и дате
        if key not in file_groups:
            file_groups[key] = []
        file_groups[key].append(row)

    # Подсвечиваем похожие файлы в одной папке
    for group in file_groups.values():
        for i, file1 in enumerate(group):
            for j, file2 in enumerate(group):
                if i < j:
                    # Проверяем схожесть названий и дату изменения
                    if are_filenames_similar(file1[3], file2[3]) and file1[4][:10] == file2[4][:10]:
                        # Проверка только между файлами с разными расширениями
                        if (file1[3].endswith(".rvt") and file2[3].endswith(".ifc")) or \
                           (file1[3].endswith(".ifc") and file2[3].endswith(".rvt")):
                            for f in (file1, file2):
                                for row in tree.get_children():
                                    if tree.item(row, "values")[2] == f[2]:  # Сравниваем путь
                                        tree.item(row, tags="highlight")

# Функция для обновления таблицы в интерфейсе
def update_table():
    for row in tree.get_children():
        tree.delete(row)

    # Получаем все данные из базы и сортируем их
    rows = cursor.execute("SELECT filename, parent_folder, path, last_modified, created_by FROM files").fetchall()

    # Сортируем сначала по папке, затем по имени, а потом по дате изменения
    rows = sorted(rows, key=lambda x: (x[1], x[3], x[0]))

    for row in rows:
        filename, parent_folder, path, last_modified, created_by = row
        # Определение цветовой метки по расширению
        tag = 'rvt' if filename.endswith('.rvt') else 'ifc'
        tree.insert('', 'end', values=row, tags=(tag,))
    apply_highlighting()

# Рекурсивная функция для поиска всех папок "Работа" и добавления файлов в базу
def scan_work_folders(root_folder):
    for dirpath, dirnames, filenames in os.walk(root_folder):
        if os.path.basename(dirpath) == "Работа":
            parent_folder = os.path.basename(os.path.dirname(dirpath))
            for filename in filenames:
                if filename.endswith(('.rvt', '.ifc')):
                    file_path = os.path.join(dirpath, filename)
                    last_modified = datetime.fromtimestamp(os.path.getmtime(file_path)).strftime('%Y-%m-%d %H:%M:%S')
                    created_by = getpass.getuser()
                    update_file_in_db(parent_folder, file_path, filename, last_modified, created_by)

# Класс для обработки событий изменения файлов
class FileMonitorHandler(FileSystemEventHandler):
    def process_file(self, event):
        if not event.is_directory:
            filename = os.path.basename(event.src_path)
            parent_folder = os.path.basename(os.path.dirname(os.path.dirname(event.src_path)))
            if filename.endswith(('.rvt', '.ifc')) and "Работа" in os.path.dirname(event.src_path):
                last_modified = datetime.fromtimestamp(os.path.getmtime(event.src_path)).strftime('%Y-%m-%d %H:%M:%S')
                created_by = getpass.getuser()
                update_file_in_db(parent_folder, event.src_path, filename, last_modified, created_by)

    def on_modified(self, event):
        self.process_file(event)

    def on_created(self, event):
        self.process_file(event)

    def on_deleted(self, event):
        if not event.is_directory:
            cursor.execute("DELETE FROM files WHERE path = ?", (event.src_path,))
            conn.commit()
            update_table()

# Выбор папки для мониторинга
def select_folder():
    selected_folder = filedialog.askdirectory()
    if selected_folder:
        folder_path.delete(0, 'end')
        folder_path.insert(0, selected_folder)

# Выбор пути для SQLite базы
def select_db_path():
    selected_db = filedialog.asksaveasfilename(defaultextension=".db", filetypes=[("SQLite Database", "*.db")])
    if selected_db:
        db_path.delete(0, 'end')
        db_path.insert(0, selected_db)
        connect_to_database(selected_db)

# Запуск мониторинга папки и предварительное сканирование
def start_monitoring():
    path = folder_path.get()
    if not os.path.isdir(path):
        messagebox.showerror("Error", "Invalid folder path")
        return
    if not os.path.isfile(db_path.get()):
        messagebox.showerror("Error", "Database path is not set or invalid")
        return
    scan_work_folders(path)
    observer = Observer()
    event_handler = FileMonitorHandler()
    observer.schedule(event_handler, path, recursive=True)
    observer.start()
    messagebox.showinfo("Info", f"Monitoring started in folder: {path}")
    root.after(100, update_table)

# Настройка интерфейса
root = Tk()
root.title("File Monitor")
root.geometry("800x600")

# Поля ввода и кнопки выбора
Label(root, text="Folder to Monitor:").grid(row=0, column=0, padx=10, pady=10, sticky="e")
folder_path = Entry(root, width=50)
folder_path.grid(row=0, column=1, padx=10, pady=10)
Button(root, text="Select Folder", command=select_folder).grid(row=0, column=2, padx=10, pady=10)

Label(root, text="SQLite Database Path:").grid(row=1, column=0, padx=10, pady=10, sticky="e")
db_path = Entry(root, width=50)
db_path.grid(row=1, column=1, padx=10, pady=10)
Button(root, text="Select DB Path", command=select_db_path).grid(row=1, column=2, padx=10, pady=10)

Button(root, text="Start Monitoring", command=start_monitoring).grid(row=2, column=1, columnspan=2, pady=10)

# Таблица для отображения файлов
tree = ttk.Treeview(root, columns=("Filename", "Parent Folder", "Path", "Last Modified", "Created By"), show="headings")
tree.heading("Filename", text="Filename")
tree.heading("Parent Folder", text="Parent Folder")
tree.heading("Path", text="Path")
tree.heading("Last Modified", text="Last Modified")
tree.heading("Created By", text="Created By")
tree.grid(row=3, column=0, columnspan=3, padx=10, pady=20, sticky="nsew")

root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(1, weight=1)

tree.tag_configure("highlight", background="lightgreen")
tree.tag_configure("rvt", background="lightblue")
tree.tag_configure("ifc", background="lightyellow")

root.mainloop()

# Закрытие базы данных при выходе
if conn:
    conn.close()
