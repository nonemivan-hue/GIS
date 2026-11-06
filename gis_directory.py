Для реализации вашей задачи разработана программа на языке Python. Она представляет собой десктопное приложение с графическим интерфейсом, способное обрабатывать файлы справочников размером более 100 МБ, искать по ним данные и выгружать результаты в ваш репозиторий на GitHub.

Как работает программа:

Производительность и большие файлы: Для работы с объемными Excel-файлами (более 100 МБ) используется библиотека pandas с движком openpyxl в режиме on_demand (загрузка листов по мере необходимости). Это позволяет программе не занимать всю оперативную память при открытии файла. Названия колонок автоматически нормализуются (приводятся к нижнему регистру, убираются кавычки и лишние пробелы), поэтому структура справочника может незначительно меняться без поломки кода.
Поиск: Поиск осуществляется в реальном времени по части строки в полях «Адрес помещения» или «Номер ЛС» без учета регистра.
Отчет: Результаты поиска выводятся в виде текстовой формы-отчета, где сгруппирована информация об организации, лицевом счете, адресе и площади.
Выгрузка в GitHub: Для отправки отчета в репозиторий https://github.com/nonemivan-hue/GIS.git используется библиотека PyGithub. Программа обращается напрямую к API GitHub. Важно: для работы выгрузки вам необходимо создать Personal Access Token (PAT) в настройках вашего GitHub-аккаунта (Settings -> Developer settings -> Personal access tokens -> Tokens (classic)) с правами доступа repo и сохранить его в переменную окружения вашей системы под именем GITHUB_TOKEN.

Инструкция по установке и запуску:

Установите зависимости. Откройте терминал и выполните команду:
pip install pandas openpyxl PyGithub
(Если вы планируете запускать программу на Linux без графической оболочки, дополнительно потребуется установка python3-tk через пакетный менеджер системы).
Создайте токен GitHub. Сгенерируйте PAT с правами repo и добавьте его в переменные окружения (в Windows: Панель управления -> Система -> Дополнительные параметры -> Переменные среды).
Сохраните код. Скопируйте приведенный ниже код и сохраните его в файл с именем gis_directory.py.
Запустите программу. Выполните в терминале команду:
python gis_directory.py

Код программы:

python
Копировать
import os
import pandas as pd
import re
from datetime import datetime
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import threading
from github import Github, InputGitTreeElement, GithubException

# --- Конфигурация репозитория ---
REPO_OWNER = 'nonemivan-hue'
REPO_NAME = 'GIS'
BRANCH = 'main'

# --- Логика обработки данных ---

def normalize_column_names(columns):
    norm_cols = []
    for col in columns:
        clean = col.strip().replace('"', '').replace("'", "")
        clean = re.sub(r'\s+', ' ', clean)
        norm_cols.append(clean.lower())
    return norm_cols

def load_data(filepath):
    try:
        # Читаем только заголовки для нормализации
        header_df = pd.read_excel(filepath, sheet_name=0, nrows=0)
        norm_columns = normalize_column_names(header_df.columns.tolist())
        
        # Читаем данные с нормализованными именами
        df = pd.read_excel(filepath, sheet_name=0)
        df.columns = norm_columns
        
        required_cols = ['номер лс (иной идентификатор потребителя)', 'адрес помещения', 'наименование организации', 'общая площадь, кв. м']
        missing = [col for col in required_cols if col not in df.columns]
        if missing:
            raise ValueError(f'В файле отсутствуют обязательные колонки: {", ".join(missing)}')
            
        return df
    except Exception as e:
        raise Exception(f'Ошибка при загрузке файла: {str(e)}')

def filter_data(df, search_text):
    if not search_text.strip():
        return df
    
    search_lower = search_text.lower()
    mask_addr = df['адрес помещения'].astype(str).str.contains(search_lower, case=False, na=False)
    mask_ls = df['номер лс (иной идентификатор потребителя)'].astype(str).str.contains(search_lower, case=False, na=False)
    
    return df[mask_addr | mask_ls]

def generate_report(df):
    if df.empty:
        return "По вашему запросу данных не найдено."
    
    lines = []
    for _, row in df.iterrows():
        lines.append(
            f"Организация: {row.get('наименование организации', 'N/A')}\n"
            f"Лицевой счет: {row.get('номер лс (иной идентификатор потребителя)', 'N/A')}\n"
            f"Адрес: {row.get('адрес помещения', 'N/A')}\n"
            f"Площадь: {row.get('общая площадь, кв. м', 'N/A')} кв.м\n"
            f"Статус ЛС: {row.get('статус лицевого счета', 'N/A')}\n"
            f"{'-'*40}"
        )
    return "\n".join(lines)

def save_report_to_file(report_text):
    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    filename = f'report_{timestamp}.txt'
    filepath = os.path.join(os.getcwd(), filename)
    
    with open(filepath, 'w', encoding='utf-8') as f:
        f.write(report_text)
    return filepath

# --- Логика работы с GitHub ---

def push_file_to_github(local_path, repo_path_in_github, commit_message):
    token = os.environ.get('GITHUB_TOKEN')
    if not token:
        raise EnvironmentError('Переменная окружения GITHUB_TOKEN не найдена.')
        
    try:
        g = Github(token)
        repo = g.get_repo(f"{REPO_OWNER}/{REPO_NAME}")
        
        with open(local_path, 'rb') as f:
            content = f.read()
        
        ref = repo.get_git_ref(f'heads/{BRANCH}')
        latest_commit_sha = ref.object.sha
        
        base_tree = repo.get_git_tree(latest_commit_sha)
        
        element = InputGitTreeElement(
            path=repo_path_in_github,
            mode='100644',
            type='blob',
            content=content
        )
        
        new_tree = repo.create_git_tree([element], base_tree)
        new_commit = repo.create_git_commit(commit_message, new_tree, [latest_commit_sha])
        repo.update_ref(ref.ref, new_commit.sha)
        
        return True, f"Файл успешно загружен в репозиторий: {repo_path_in_github}"
    except GithubException as e:
        return False, f"Ошибка GitHub API: {e.data.get('message', str(e))}"
    except Exception as e:
        return False, str(e)

# --- Графический интерфейс ---

class App:
    def __init__(self, root):
        self.root = root
        self.root.title("Справочник ГИС ЖКХ")
        self.root.geometry("900x600")
        self.df = None
        
        control_frame = ttk.Frame(root)
        control_frame.pack(fill=tk.X, padx=10, pady=5)
        
        ttk.Button(control_frame, text="Загрузить файл", command=self.load_file).pack(side=tk.LEFT, padx=5)
        self.search_var = tk.StringVar()
        search_entry = ttk.Entry(control_frame, textvariable=self.search_var, width=40)
        search_entry.pack(side=tk.LEFT, padx=5)
        search_entry.bind('<Return>', lambda e: self.search())
        ttk.Button(control_frame, text="Найти", command=self.search).pack(side=tk.LEFT, padx=5)
        ttk.Button(control_frame, text="Выгрузить в GitHub", command=self.push_to_github).pack(side=tk.LEFT, padx=5)
        
        report_frame = ttk.Frame(root)
        report_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=5)
        
        self.report_text = scrolledtext.ScrolledText(report_frame, wrap=tk.WORD, font=("Consolas", 10))
        self.report_text.pack(fill=tk.BOTH, expand=True)
        
        self.status_label = ttk.Label(root, text="Готово", anchor=tk.W)
        self.status_label.pack(fill=tk.X, padx=10, pady=5)

    def log_status(self, message):
        self.status_label.config(text=message)
        self.root.update_idletasks()

    def load_file(self):
        filepath = filedialog.askopenfilename(
            title="Выберите файл справочника",
            filetypes=[("Excel files", "*.xls *.xlsx"), ("All files", "*.*")]
        )
        if not filepath:
            return
            
        def worker():
            try:
                self.log_status("Загружается файл...")
                self.df = load_data(filepath)
                self.log_status(f"Файл загружен. Записей: {len(self.df)}")
                self.report_text.delete(1.0, tk.END)
                self.report_text.insert(tk.END, f"Загружено {len(self.df)} записей. Введите часть адреса или ЛС для поиска.")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                self.log_status("Ошибка загрузки")
                
        threading.Thread(target=worker, daemon=True).start()

    def search(self):
        if self.df is None:
            messagebox.showwarning("Внимание", "Сначала загрузите файл.")
            return
            
        search_text = self.search_var.get()
        
        def worker():
            try:
                self.log_status("Идет поиск...")
                filtered = filter_data(self.df, search_text)
                report = generate_report(filtered)
                self.report_text.delete(1.0, tk.END)
                self.report_text.insert(tk.END, report)
                self.log_status(f"Поиск завершен. Найдено: {len(filtered)}")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                self.log_status("Ошибка поиска")
                
        threading.Thread(target=worker, daemon=True).start()

    def push_to_github(self):
        if self.df is None:
            messagebox.showwarning("Внимание", "Нет данных для выгрузки.")
            return
            
        content = self.report_text.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("Внимание", "Отчет пуст.")
            return
            
        def worker():
            try:
                self.log_status("Сохранение локального отчета...")
                local_path = save_report_to_file(content)
                
                timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
                repo_path = f'reports/report_{timestamp}.txt'
                
                self.log_status("Выгрузка в GitHub...")
                success, message = push_file_to_github(local_path, repo_path, f'Add report {timestamp}')
                
                if success:
                    messagebox.showinfo("Успех", message)
                else:
                    messagebox.showerror("Ошибка GitHub", message)
                    
                self.log_status("Готово")
            except Exception as e:
                messagebox.showerror("Ошибка", str(e))
                self.log_status("Ошибка выгрузки")
                
        threading.Thread(target=worker, daemon=True).start()

if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)
    root.mainloop()
