import configparser
import requests
import json
import base64
from datetime import datetime, timezone, timedelta
import openpyxl
from collections import defaultdict
import tkinter as tk
from tkinter import ttk, filedialog
import os
import sys
import platform

class ZoomWebinarApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Zoom Webinar Report Generator")
        self.root.geometry("600x800")
        self.root.configure(bg="#1C2526")
        self.root.minsize(500, 600)

        self.account_id = tk.StringVar()
        self.client_id = tk.StringVar()
        self.client_secret = tk.StringVar()
        self.user_id = tk.StringVar()
        self.session_count = tk.StringVar(value="10")
        self.save_dir = tk.StringVar(value=os.getcwd())

        self.main_frame = tk.Frame(self.root, bg="#1C2526", padx=10, pady=10)
        self.main_frame.grid(row=0, column=0, sticky="nsew")

        self.bg_color = "#1C2526"
        self.fg_color = "#E8ECEF"
        self.entry_bg = "#2D3B3D"
        self.button_bg = "#3B4A4C"
        self.button_active_bg = "#5A6A6C"
        self.listbox_bg = "#2D3B3D"
        self.listbox_select_bg = "#5A6A6C"

        tk.Label(self.main_frame, text="Account ID:", bg=self.bg_color, fg=self.fg_color).grid(row=0, column=0, sticky="w", pady=5)
        tk.Entry(self.main_frame, textvariable=self.account_id, width=40, bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color).grid(row=0, column=1, sticky="ew", pady=5)

        tk.Label(self.main_frame, text="Client ID:", bg=self.bg_color, fg=self.fg_color).grid(row=1, column=0, sticky="w", pady=5)
        tk.Entry(self.main_frame, textvariable=self.client_id, width=40, bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color).grid(row=1, column=1, sticky="ew", pady=5)

        tk.Label(self.main_frame, text="Client Secret:", bg=self.bg_color, fg=self.fg_color).grid(row=2, column=0, sticky="w", pady=5)
        tk.Entry(self.main_frame, textvariable=self.client_secret, width=40, show="*", bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color).grid(row=2, column=1, sticky="ew", pady=5)

        tk.Label(self.main_frame, text="User ID:", bg=self.bg_color, fg=self.fg_color).grid(row=3, column=0, sticky="w", pady=5)
        tk.Entry(self.main_frame, textvariable=self.user_id, width=40, bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color).grid(row=3, column=1, sticky="ew", pady=5)

        tk.Label(self.main_frame, text="Количество сессий:", bg=self.bg_color, fg=self.fg_color).grid(row=4, column=0, sticky="w", pady=5)
        self.session_combo = ttk.Combobox(self.main_frame, textvariable=self.session_count, values=["10", "20", "30"], state="readonly")
        self.session_combo.grid(row=4, column=1, sticky="w", pady=5)
        self.session_combo.configure(background=self.entry_bg, foreground=self.fg_color)

        tk.Label(self.main_frame, text="Директория сохранения:", bg=self.bg_color, fg=self.fg_color).grid(row=5, column=0, sticky="w", pady=5)
        tk.Entry(self.main_frame, textvariable=self.save_dir, width=30, bg=self.entry_bg, fg=self.fg_color, insertbackground=self.fg_color).grid(row=5, column=1, sticky="w", pady=5)
        tk.Button(self.main_frame, text="Выбрать", command=self.choose_directory, bg=self.button_bg, fg=self.fg_color, activebackground=self.button_active_bg).grid(row=5, column=1, sticky="e", pady=5)

        tk.Button(self.main_frame, text="Загрузить вебинары", command=self.load_webinars, bg=self.button_bg, fg=self.fg_color, activebackground=self.button_active_bg).grid(row=6, column=0, columnspan=2, pady=10, sticky="ew")

        tk.Label(self.main_frame, text="Выберите вебинары:", bg=self.bg_color, fg=self.fg_color).grid(row=7, column=0, sticky="w", pady=5)
        self.webinar_listbox = tk.Listbox(self.main_frame, selectmode="multiple", width=50, height=12, bg=self.listbox_bg, fg=self.fg_color, selectbackground=self.listbox_select_bg, selectforeground=self.fg_color)
        self.webinar_listbox.grid(row=8, column=0, columnspan=2, pady=5, sticky="nsew")

        tk.Button(self.main_frame, text="Обработать выбранные вебинары", command=self.process_selected_webinars, bg=self.button_bg, fg=self.fg_color, activebackground=self.button_active_bg).grid(row=9, column=0, columnspan=2, pady=10, sticky="ew")

        self.progress = ttk.Progressbar(self.main_frame, mode="determinate", length=300)
        self.progress.grid(row=10, column=0, columnspan=2, pady=5, sticky="ew")

        self.log_text = tk.Text(self.main_frame, height=15, bg=self.listbox_bg, fg=self.fg_color, wrap="word", insertbackground=self.fg_color, state="normal")
        self.log_text.grid(row=11, column=0, columnspan=2, sticky="nsew", pady=5)
        self.log_text.insert(tk.END, "Логи будут отображаться здесь...\n")
        self.log_text.bind("<Key>", lambda e: "break")  # Блокировать редактирование
        self.log_text.see(tk.END)

        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        self.main_frame.columnconfigure(1, weight=1)
        self.main_frame.rowconfigure(8, weight=2)
        self.main_frame.rowconfigure(11, weight=3)

        self.all_sessions = []

        self.log(f"Python version: {platform.python_version()}")
        self.log(f"Script path: {os.path.abspath(__file__)}")

        self.load_config()

    def log(self, message):
        if hasattr(self, 'log_text'):
            self.log_text.insert(tk.END, f"{message}\n")
            self.log_text.see(tk.END)
        else:
            print(f"LOG: {message}")

    def load_config(self):
        config = configparser.ConfigParser()
        try:
            if os.path.exists("config.ini"):
                config.read("config.ini")
                if "ZoomCredentials" in config:
                    self.account_id.set(config["ZoomCredentials"].get("account_id", ""))
                    self.client_id.set(config["ZoomCredentials"].get("client_id", ""))
                    self.client_secret.set(config["ZoomCredentials"].get("client_secret", ""))
                    self.user_id.set(config["ZoomCredentials"].get("user_id", ""))
                    self.log("Настройки Zoom загружены из config.ini")
                else:
                    self.log("Секция ZoomCredentials не найдена в config.ini")
            else:
                self.log("Файл config.ini не найден, используйте поля для ввода настроек")
        except Exception as e:
            self.log(f"Ошибка при загрузке config.ini: {str(e)}")

    def save_config(self):
        config = configparser.ConfigParser()
        try:
            config["ZoomCredentials"] = {
                "account_id": self.account_id.get(),
                "client_id": self.client_id.get(),
                "client_secret": self.client_secret.get(),
                "user_id": self.user_id.get()
            }
            with open("config.ini", "w", encoding="utf-8") as configfile:
                config.write(configfile)
            self.log("Настройки Zoom сохранены в config.ini")
        except Exception as e:
            self.log(f"Ошибка при сохранении config.ini: {str(e)}")

    def choose_directory(self):
        directory = filedialog.askdirectory(initialdir=self.save_dir.get())
        if directory:
            self.save_dir.set(directory)
            self.log(f"Выбрана директория: {directory}")

    def get_access_token(self):
        auth_string = f"{self.client_id.get()}:{self.client_secret.get()}"
        auth_bytes = auth_string.encode('ascii')
        auth_base64 = base64.b64encode(auth_bytes).decode('ascii')
        
        token_url = "https://zoom.us/oauth/token"
        headers = {
            "Authorization": f"Basic {auth_base64}",
            "Content-Type": "application/x-www-form-urlencoded"
        }
        payload = {
            "grant_type": "account_credentials",
            "account_id": self.account_id.get()
        }
        
        response = requests.post(token_url, headers=headers, data=payload)
        if response.status_code == 200:
            return response.json()['access_token']
        else:
            self.log(f"Ошибка получения токена: {response.status_code} - {response.text}")
            raise Exception("Ошибка получения токена")

    def get_past_webinars(self, access_token):
        webinars_url = f"https://api.zoom.us/v2/users/{self.user_id.get()}/webinars"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        params = {
            "page_size": int(self.session_count.get()),
            "type": "past"
        }
        
        response = requests.get(webinars_url, headers=headers, params=params)
        if response.status_code == 200:
            return response.json().get("webinars", [])
        else:
            self.log(f"Ошибка получения списка вебинаров: {response.status_code} - {response.text}")
            raise Exception("Ошибка получения списка вебинаров")

    def get_webinar_instances(self, access_token, webinar_id):
        instances_url = f"https://api.zoom.us/v2/past_webinars/{webinar_id}/instances"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(instances_url, headers=headers)
        if response.status_code == 200:
            return response.json().get("webinars", [])
        else:
            self.log(f"Ошибка получения сессий вебинара {webinar_id}: {response.status_code} - {response.text}")
            raise Exception(f"Ошибка получения сессий вебинара {webinar_id}")

    def get_webinar_participants(self, access_token, webinar_uuid):
        report_url = f"https://api.zoom.us/v2/report/webinars/{webinar_uuid}/participants"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        params = {
            "page_size": 300
        }
        
        participants = []
        next_page_token = ""
        
        while True:
            if next_page_token:
                params["next_page_token"] = next_page_token
            response = requests.get(report_url, headers=headers, params=params)
            if response.status_code == 200:
                data = response.json()
                participants.extend(data.get("participants", []))
                next_page_token = data.get("next_page_token", "")
                if not next_page_token:
                    break
            else:
                self.log(f"Ошибка получения отчёта участников: {response.status_code} - {response.text}")
                raise Exception("Ошибка получения отчёта участников")
        
        return participants

    def get_webinar_panelists(self, access_token, webinar_id):
        panelists_url = f"https://api.zoom.us/v2/webinars/{webinar_id}/panelists"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(panelists_url, headers=headers)
        if response.status_code == 200:
            return response.json().get("panelists", [])
        else:
            self.log(f"Ошибка получения данных панелистов: {response.status_code} - {response.text}")
            raise Exception("Ошибка получения данных панелистов")

    def get_webinar_registrants(self, access_token, webinar_id):
        registrants_url = f"https://api.zoom.us/v2/webinars/{webinar_id}/registrants"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        params = {
            "page_size": 300,
            "status": "approved"
        }
        
        registrants = []
        next_page_token = ""
        
        while True:
            if next_page_token:
                params["next_page_token"] = next_page_token
            response = requests.get(registrants_url, headers=headers, params=params)
            if response.status_code == 200:
                data = response.json()
                registrants.extend(data.get("registrants", []))
                next_page_token = data.get("next_page_token", "")
                if not next_page_token:
                    break
            else:
                self.log(f"Ошибка получения данных регистраций: {response.status_code} - {response.text}")
                raise Exception("Ошибка получения данных регистраций")
        
        return registrants

    def get_webinar_polls(self, access_token, webinar_uuid):
        polls_url = f"https://api.zoom.us/v2/report/webinars/{webinar_uuid}/polls"
        headers = {
            "Authorization": f"Bearer {access_token}",
            "Content-Type": "application/json"
        }
        
        response = requests.get(polls_url, headers=headers)
        if response.status_code == 200:
            return response.json().get("questions", [])
        elif response.status_code == 404:
            return []
        else:
            self.log(f"Ошибка получения данных опросов: {response.status_code} - {response.text}")
            raise Exception("Ошибка получения данных опросов")

    def get_custom_question_value(self, custom_questions, title):
        if not custom_questions:
            return ""
        for question in custom_questions:
            if question.get("title") == title:
                return question.get("value", "")
        return ""

    def extract_time(self, time_str, add_hours=0):
        if not time_str:
            return ""
        try:
            if "T" in time_str and "Z" in time_str:
                dt = datetime.fromisoformat(time_str.replace("Z", "+00:00"))
                dt = dt + timedelta(hours=add_hours)
                return dt.strftime("%H:%M:%S")
            else:
                dt = datetime.strptime(time_str, "%Y-%m-%d %H:%M:%S")
                dt = dt + timedelta(hours=add_hours)
                return dt.strftime("%H:%M:%S")
        except:
            return ""

    def extract_date(self, session_date):
        try:
            dt = datetime.fromisoformat(session_date.replace("Z", "+00:00"))
            return dt.strftime("%d.%m.%Y")
        except:
            return session_date[:10].replace("-", ".")

    def is_valid_int(self, value):
        if not value:
            return False
        try:
            int(value)
            return True
        except ValueError:
            return False

    def merge_participant_data(self, participants, registrants, panelists):
        merged_data = []
        unmatched_participants = 0
        email_to_sessions = defaultdict(list)
        
        panelist_dict = {p.get("email", "").lower(): p for p in panelists if p.get("email")}
        
        if participants and all(not p.get("user_email") and not p.get("registrant_id") for p in participants):
            self.log("ВНИМАНИЕ: Данные участников ограничены. Используем данные регистрантов и панелистов.")
            
            for registrant in registrants:
                custom_questions = registrant.get("custom_questions", [])
                participant_data = {
                    "Имя пользователя (исходное имя)": registrant.get("first_name", "") + " " + registrant.get("last_name", ""),
                    "Имя": registrant.get("first_name", ""),
                    "Фамилия": registrant.get("last_name", ""),
                    "Эл. почта": registrant.get("email", ""),
                    "Город": registrant.get("city", ""),
                    "Телефон": registrant.get("phone", ""),
                    "Организация": self.get_custom_question_value(custom_questions, "Организация"),
                    "Должность": registrant.get("job_title", ""),
                    "Специальность": self.get_custom_question_value(custom_questions, "Специальность"),
                    "Время регистрации": registrant.get("create_time", ""),
                    "Время входа": "",
                    "Время выхода": "",
                    "Время в сеансе (минут)": 0,
                    "Является гостем": False,
                    "Название страны/региона": registrant.get("country", ""),
                    "Роль": "attendee"
                }
                merged_data.append(participant_data)
            
            for participant in participants:
                role = participant.get("role", "").lower()
                if role in ["host", "panelist"]:
                    participant_data = {
                        "Имя пользователя (исходное имя)": participant.get("name", ""),
                        "Эл. почта": participant.get("user_email", ""),
                        "Город": "",
                        "Телефон": "",
                        "Организация": "",
                        "Должность": "",
                        "Специальность": "",
                        "Время регистрации": "",
                        "Состояние утверждения": "",
                        "Время входа": self.extract_time(participant.get("join_time", ""), add_hours=3),
                        "Время выхода": self.extract_time(participant.get("leave_time", ""), add_hours=3),
                        "Время в сеансе (минут)": 0,
                        "Является гостем": participant.get("is_guest", False),
                        "Название страны/региона": participant.get("country", ""),
                        "Роль": role
                    }
                    name = participant.get("name", "")
                    if name == "Pain Russia":
                        participant_data["Имя"] = "ROIB"
                        participant_data["Фамилия"] = "PainRussia"
                    else:
                        name_parts = name.split(" ", 1)
                        participant_data["Фамилия"] = name_parts[0] if name_parts else ""
                        participant_data["Имя"] = name_parts[1] if len(name_parts) > 1 else ""
                    duration = participant.get("duration")
                    if duration:
                        participant_data["Время в сеансе (минут)"] = round(duration / 60, 2)
                    else:
                        try:
                            join_time = datetime.fromisoformat(participant.get("join_time", "").replace("Z", "+00:00"))
                            leave_time = datetime.fromisoformat(participant.get("leave_time", "").replace("Z", "+00:00"))
                            duration = (leave_time - join_time).total_seconds() / 60
                            participant_data["Время в сеансе (минут)"] = round(duration, 2) if duration > 0 else 0
                        except:
                            participant_data["Время в сеансе (минут)"] = 0
                    email = participant_data["Эл. почта"].lower()
                    if email:
                        email_to_sessions[email].append(participant_data)
                    merged_data.append(participant_data)
            
            for panelist in panelists:
                email = panelist.get("email", "").lower()
                if email and not any(p["Эл. почта"].lower() == email for p in merged_data):
                    participant_data = {
                        "Имя пользователя (исходное имя)": panelist.get("name", ""),
                        "Эл. почта": panelist.get("email", ""),
                        "Город": "",
                        "Телефон": "",
                        "Организация": "",
                        "Должность": "",
                        "Специальность": "",
                        "Время регистрации": "",
                        "Состояние утверждения": "",
                        "Время входа": "",
                        "Время выхода": "",
                        "Время в сеансе (минут)": 0,
                        "Является гостем": False,
                        "Название страны/региона": "",
                        "Роль": "panelist"
                    }
                    name = panelist.get("name", "")
                    if name == "Pain Russia":
                        participant_data["Имя"] = "ROIB"
                        participant_data["Фамилия"] = "PainRussia"
                    else:
                        name_parts = name.split(" ", 1)
                        participant_data["Фамилия"] = name_parts[0] if name_parts else ""
                        participant_data["Имя"] = name_parts[1] if len(name_parts) > 1 else ""
                    merged_data.append(participant_data)
            
            return merged_data, email_to_sessions
        
        for participant in participants:
            role = participant.get("role", "attendee").lower()
            name = participant.get("name", "")
            participant_data = {
                "Имя пользователя (исходное имя)": name,
                "Эл. почта": participant.get("user_email", ""),
                "Город": "",
                "Телефон": "",
                "Организация": "",
                "Должность": "",
                "Специальность": "",
                "Время регистрации": "",
                "Состояние утверждения": "",
                "Время входа": self.extract_time(participant.get("join_time", ""), add_hours=3),
                "Время выхода": self.extract_time(participant.get("leave_time", ""), add_hours=3),
                "Время в сеансе (минут)": 0,
                "Является гостем": participant.get("is_guest", False),
                "Название страны/региона": participant.get("country", ""),
                "Роль": role
            }
            if name == "Pain Russia":
                participant_data["Имя"] = "ROIB"
                participant_data["Фамилия"] = "PainRussia"
            elif role == "panelist":
                name_parts = name.split(" ", 1)
                participant_data["Фамилия"] = name_parts[0] if name_parts else ""
                participant_data["Имя"] = name_parts[1] if len(name_parts) > 1 else ""
            else:
                participant_data["Имя"] = participant.get("first_name", "")
                participant_data["Фамилия"] = participant.get("last_name", "")
            
            duration = participant.get("duration")
            if duration:
                participant_data["Время в сеансе (минут)"] = round(duration / 60, 2)
            else:
                try:
                    join_time = datetime.fromisoformat(participant.get("join_time", "").replace("Z", "+00:00"))
                    leave_time = datetime.fromisoformat(participant.get("leave_time", "").replace("Z", "+00:00"))
                    duration = (leave_time - join_time).total_seconds() / 60
                    participant_data["Время в сеансе (минут)"] = round(duration, 2) if duration > 0 else 0
                except:
                    participant_data["Время в сеансе (минут)"] = 0
            
            matched = False
            participant_email = participant.get("user_email", "").lower()
            participant_reg_id = participant.get("registrant_id", "")
            
            for registrant in registrants:
                registrant_email = registrant.get("email", "").lower()
                registrant_id = registrant.get("id", "")
                
                if (participant_email and registrant_email and participant_email == registrant_email) or \
                   (participant_reg_id and registrant_id and participant_reg_id == registrant_id):
                    custom_questions = registrant.get("custom_questions", [])
                    update_data = {
                        "Город": registrant.get("city", ""),
                        "Телефон": registrant.get("phone", ""),
                        "Организация": self.get_custom_question_value(custom_questions, "Организация"),
                        "Должность": registrant.get("job_title", ""),
                        "Специальность": self.get_custom_question_value(custom_questions, "Специальность"),
                        "Время регистрации": registrant.get("create_time", ""),
                        "Эл. почта": participant_data["Эл. почта"] or registrant.get("email", ""),
                        "Название страны/региона": participant_data["Название страны/региона"] or registrant.get("country", "")
                    }
                    if role != "panelist" and participant_data["Имя пользователя (исходное имя)"] != "Pain Russia":
                        update_data["Имя"] = registrant.get("first_name", "")
                        update_data["Фамилия"] = registrant.get("last_name", "")
                    participant_data.update(update_data)
                    matched = True
                    break
            
            if participant_email in panelist_dict:
                update_data = {
                    "Имя пользователя (исходное имя)": participant_data["Имя пользователя (исходное имя)"] or panelist_dict[participant_email].get("name", ""),
                    "Эл. почта": participant_data["Эл. почта"] or panelist_dict[participant_email].get("email", ""),
                    "Роль": "panelist"
                }
                name = panelist_dict[participant_email].get("name", "")
                if name == "Pain Russia":
                    update_data["Имя"] = "ROIB"
                    update_data["Фамилия"] = "PainRussia"
                else:
                    name_parts = name.split(" ", 1)
                    update_data["Фамилия"] = name_parts[0] if name_parts else ""
                    update_data["Имя"] = name_parts[1] if len(name_parts) > 1 else ""
                participant_data.update(update_data)
                matched = True
            
            if not matched:
                unmatched_participants += 1
            
            email = participant_data["Эл. почта"].lower()
            if email:
                email_to_sessions[email].append(participant_data)
            merged_data.append(participant_data)
        
        if unmatched_participants > 0:
            self.log(f"ВНИМАНИЕ: {unmatched_participants} участников не сопоставлены с регистрантами или панелистами.")
        
        merged_data.sort(key=lambda x: x["Роль"] != "panelist")
        
        return merged_data, email_to_sessions

    def process_polls(self, polls):
        poll_times = {}
        max_polls = 0
        for participant in polls:
            email = participant.get("email", "").lower()
            if not email:
                continue
            question_details = participant.get("question_details", [])
            max_polls = max(max_polls, len(question_details))
            question_details = sorted(question_details, key=lambda x: x.get("date_time", ""), reverse=False)[:10]
            for idx, detail in enumerate(question_details, 1):
                date_time = detail.get("date_time", "")
                if date_time:
                    poll_times.setdefault(email, {})[f"Опрос {idx}"] = self.extract_time(date_time, add_hours=3)
        return poll_times, min(max_polls, 10)

    def save_to_excel(self, participants, email_to_sessions, polls, session_date, webinar_id):
        wb_zoom = openpyxl.Workbook()
        ws_zoom = wb_zoom.active
        ws_zoom.title = "Zoom Report"
        
        wb_roib = openpyxl.Workbook()
        ws_roib = wb_roib.active
        ws_roib.title = "ROIB Report"
        
        headers_zoom = [
            "Имя пользователя (исходное имя)", "Имя", "Фамилия", "Эл. почта",
            "Город", "Телефон", "Организация", "Должность", "Специальность",
            "Время регистрации", "Время входа", "Время выхода", "Время в сеансе (минут)",
            "Название страны/региона", "Роль"
        ]
        headers_roib = [
            "Время в сеансе (минут)", "Фамилия", "Имя", "Город", "Организация",
            "Должность", "Специальность", "Дата"
        ]
        
        poll_times, num_polls = self.process_polls(polls)
        
        for i in range(1, num_polls + 1):
            headers_zoom.append(f"Опрос {i}")
            headers_roib.append(f"Опрос {i}")
        
        ws_zoom.append(headers_zoom)
        ws_roib.append(headers_roib)
        
        event_date_str = self.extract_date(session_date)
        
        processed_emails = set()
        for participant in participants:
            email = participant["Эл. почта"].lower()
            if email in processed_emails:
                continue
            
            total_duration = sum(p["Время в сеансе (минут)"] for p in email_to_sessions.get(email, [participant]))
            participant["Время в сеансе (минут)"] = round(total_duration, 2)
            
            poll_data = poll_times.get(email, {})
            
            row_zoom = [
                participant.get("Имя пользователя (исходное имя)", ""),
                participant.get("Имя", ""),
                participant.get("Фамилия", ""),
                participant.get("Эл. почта", ""),
                participant.get("Город", ""),
                participant.get("Телефон", ""),
                participant.get("Организация", ""),
                participant.get("Должность", ""),
                participant.get("Специальность", ""),
                participant.get("Время регистрации", ""),
                participant.get("Время входа", ""),
                participant.get("Время выхода", ""),
                participant.get("Время в сеансе (минут)", 0),
                participant.get("Название страны/региона", ""),
                participant.get("Роль", "attendee")
            ]
            row_zoom.extend([poll_data.get(f"Опрос {i}", "") for i in range(1, num_polls + 1)])
            ws_zoom.append(row_zoom)
            
            row_roib = [
                participant.get("Время в сеансе (минут)", 0),
                participant.get("Фамилия", ""),
                participant.get("Имя", ""),
                participant.get("Город", ""),
                participant.get("Организация", ""),
                participant.get("Должность", ""),
                participant.get("Специальность", ""),
                event_date_str
            ]
            row_roib.extend([poll_data.get(f"Опрос {i}", "") for i in range(1, num_polls + 1)])
            ws_roib.append(row_roib)
            
            processed_emails.add(email)
        
        try:
            event_date = datetime.fromisoformat(session_date.replace("Z", "+00:00")).strftime("%y%m%d")
        except:
            event_date = session_date[:10].replace("-", "")[2:]
        
        output_dir = os.path.join(self.save_dir.get(), event_date)
        os.makedirs(output_dir, exist_ok=True)
        
        zoom_file = os.path.join(output_dir, f"Zoom{event_date}.xlsx")
        roib_file = os.path.join(output_dir, f"roib{event_date}.xlsx")
        
        wb_zoom.save(zoom_file)
        wb_roib.save(roib_file)
        self.log(f"Отчёт сохранён в: {zoom_file}")
        self.log(f"Дополнительный отчёт сохранён в: {roib_file}")

    def load_webinars(self):
        self.log("Начало загрузки вебинаров")
        try:
            if not all([self.account_id.get(), self.client_id.get(), self.client_secret.get(), self.user_id.get()]):
                self.log("Ошибка: Заполните все поля Zoom API")
                return
            
            self.webinar_listbox.delete(0, tk.END)
            self.all_sessions = []
            
            self.log("Получение токена доступа")
            access_token = self.get_access_token()
            self.log("Получение списка вебинаров")
            webinars = self.get_past_webinars(access_token)
            
            if not webinars:
                self.log("Прошедшие вебинары не найдены")
                return
            
            self.log("Получение сессий вебинаров")
            for webinar in webinars:
                webinar_id = webinar["id"]
                current_topic = webinar.get("topic", "Без названия")
                try:
                    instances = self.get_webinar_instances(access_token, webinar_id)
                    for instance in instances:
                        self.all_sessions.append({
                            "webinar_id": webinar_id,
                            "uuid": instance["uuid"],
                            "start_time": instance["start_time"],
                            "topic": current_topic
                        })
                except Exception as e:
                    self.log(f"Ошибка при получении сессий для вебинара {webinar_id}: {str(e)}")
            
            if not self.all_sessions:
                self.log("Прошедшие сессии вебинаров не найдены")
                return
            
            self.all_sessions.sort(key=lambda x: x["start_time"], reverse=True)
            
            for idx, session in enumerate(self.all_sessions, 1):
                start_time = session["start_time"]
                topic = session["topic"]
                try:
                    dt = datetime.fromisoformat(start_time.replace("Z", "+00:00")) + timedelta(hours=4)
                    formatted_date = dt.strftime("%Y-%m-%d %H:%M")
                except:
                    formatted_date = start_time
                self.webinar_listbox.insert(tk.END, f"{idx}. {formatted_date} - {topic} (Webinar ID: {session['webinar_id']})")
            
            self.log(f"Загружено {len(self.all_sessions)} сессий вебинаров")
        
        except Exception as e:
            self.log(f"Ошибка при загрузке вебинаров: {str(e)}")

    def process_selected_webinars(self):
        selected_indices = self.webinar_listbox.curselection()
        if not selected_indices:
            self.log("Ошибка: Выберите хотя бы один вебинар")
            return
        
        try:
            self.progress["value"] = 0
            self.progress["maximum"] = 100
            step = 100 / (7 * len(selected_indices))  # 7 шагов на вебинар
            
            for idx in selected_indices:
                session = self.all_sessions[idx]
                webinar_id = session["webinar_id"]
                webinar_uuid = session["uuid"]
                webinar_topic = session["topic"].replace(" ", "_").replace("/", "_")
                session_date = session["start_time"].replace(":", "-").replace("Z", "")
                
                self.log(f"Обработка вебинара: {session['topic']} (ID: {webinar_id})")
                
                self.log("Получение токена доступа")
                access_token = self.get_access_token()
                self.progress["value"] += step
                self.root.update()
                
                self.save_config()
                
                self.log(f"Получение участников вебинара ID: {webinar_id}")
                participants = self.get_webinar_participants(access_token, webinar_uuid)
                self.progress["value"] += step
                self.root.update()
                
                self.log(f"Получение регистрантов вебинара ID: {webinar_id}")
                registrants = self.get_webinar_registrants(access_token, webinar_id)
                self.progress["value"] += step
                self.root.update()
                
                self.log(f"Получение панелистов вебинара ID: {webinar_id}")
                panelists = self.get_webinar_panelists(access_token, webinar_id)
                self.progress["value"] += step
                self.root.update()
                
                self.log(f"Получение опросов вебинара ID: {webinar_id}")
                polls = self.get_webinar_polls(access_token, webinar_uuid)
                self.progress["value"] += step
                self.root.update()
                
                self.log("Объединение данных участников")
                merged_participants, email_to_sessions = self.merge_participant_data(participants, registrants, panelists)
                
                output_dir = os.path.join(self.save_dir.get(), session_date[:10].replace("-", "")[2:])
                os.makedirs(output_dir, exist_ok=True)
                
                self.log(f"Сохранение JSON отчёта участников")
                participants_json_file = os.path.join(output_dir, f"participants_{webinar_id}_{session_date}_{webinar_topic}.json")
                with open(participants_json_file, 'w', encoding='utf-8') as f:
                    json.dump(merged_participants, f, ensure_ascii=False, indent=2)
                self.log(f"Отчёт об участниках сохранён в JSON: {participants_json_file}")
                
                self.log(f"Сохранение JSON данных опросов")
                polls_file = os.path.join(output_dir, f"polls_{webinar_id}_{session_date}_{webinar_topic}.json")
                with open(polls_file, 'w', encoding='utf-8') as f:
                    json.dump(polls, f, ensure_ascii=False, indent=2)
                if polls:
                    self.log(f"Данные по опросам сохранены в {polls_file}")
                    self.log(f"Количество записей опросов: {len(polls)}")
                else:
                    self.log("Опросы для этой сессии отсутствуют")
                self.progress["value"] += step
                self.root.update()
                
                self.log(f"Сохранение Excel отчётов")
                self.save_to_excel(merged_participants, email_to_sessions, polls, session_date, webinar_id)
                self.progress["value"] += step
                self.root.update()
                
                self.log(f"Всего участников: {len(set(p['Эл. почта'].lower() for p in merged_participants if p['Эл. почта']))}")
            
            self.log("Обработка всех выбранных вебинаров завершена")
            self.progress["value"] = 100
            self.root.update()
        
        except Exception as e:
            self.log(f"Ошибка при обработке вебинаров: {str(e)}")
            self.progress["value"] = 0
            self.root.update()

if __name__ == "__main__":
    root = tk.Tk()
    app = ZoomWebinarApp(root)
    root.mainloop()
