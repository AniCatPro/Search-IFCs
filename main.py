import os
import sqlite3
from datetime import datetime
import getpass
from tkinter import Tk, filedialog, Label, Button, Entry, messagebox
from tkinter import ttk
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler


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


# Функция для проверки и подсветки файлов
def apply_highlighting():
    rows = cursor.execute("SELECT * FROM files").fetchall()
    file_groups = {}

    # Группируем файлы по папке и имени файла без расширения
    for row in rows:
        parent_folder, path, filename, last_modified, created_by = row[1:]
        base_name, ext = os.path.splitext(filename)
        key = (parent_folder, base_name)

        # Группируем файлы по папке и базовому имени
        if key not in file_groups:
            file_groups[key] = []
        file_groups[key].append(row)

    # Подсвечиваем файлы с одинаковым базовым именем и последним изменением в один и тот же день
    for files in file_groups.values():
        if len(files) > 1:
            dates = {f[4][:10] for f in files}  # Извлекаем даты (до дня)
            if len(dates) == 1:  # Если дата последнего изменения одинакова
                for f in files:
                    for row in tree.get_children():
                        if tree.item(row, "values")[2] == f[2]:  # Сравниваем путь
                            tree.item(row, tags="highlight")


# Функция для отображения таблицы и применения подсветки
def update_table():
    for row in tree.get_children():
        tree.delete(row)
    cursor.execute("SELECT filename, parent_folder, path, last_modified, created_by FROM files")
    for row in cursor.fetchall():
        tree.insert('', 'end', values=row)
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
tree.tag_configure("highlight", background="lightgreen")
tree.grid(row=3, column=0, columnspan=3, padx=10, pady=20, sticky="nsew")

root.grid_rowconfigure(3, weight=1)
root.grid_columnconfigure(1, weight=1)

root.mainloop()

if conn:
    conn.close()
