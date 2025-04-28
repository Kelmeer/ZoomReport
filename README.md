# Zoom Webinar Report Generator

![Python](https://img.shields.io/badge/Python-3.8%2B-blue)
![License](https://img.shields.io/badge/License-MIT-green)

## Описание

**Zoom Webinar Report Generator** — это Python-приложение с графическим интерфейсом (GUI) для получения и обработки данных о прошедших вебинарах Zoom через Zoom API. Программа собирает информацию об участниках, регистрантах, панелистах и опросах, объединяет данные и формирует отчёты в формате Excel (`Zoom{YYMMDD}.xlsx`, `roib{YYMMDD}.xlsx`) и JSON. Подходит для организаторов вебинаров, аналитиков и администраторов, которым нужно автоматизировать обработку данных Zoom.

### Основные возможности
- Загрузка списка прошедших вебинаров (до 10, 20 или 30 сессий).
- Получение данных об участниках, регистрантах, панелистах и опросах через Zoom API.
- Объединение данных с учётом ролей (`panelist`, `attendee`, `host`).
- Специальная обработка имён:
  - Для панелистов: `Фамилия` — первое слово из имени, `Имя` — остальное.
  - Для пользователя `"Pain Russia"`: `Имя: ROIB`, `Фамилия: PainRussia`.
  - Для остальных: `Имя` и `Фамилия` из Zoom API (`first_name`, `last_name`).
- Формирование отчётов:
  - `Zoom{YYMMDD}.xlsx`: Полные данные (имя, email, время входа/выхода, опросы и др.).
  - `roib{YYMMDD}.xlsx`: Упрощённый отчёт (время, имя, город, организация и др.).
  - JSON-файлы: Данные участников и опросов.
- GUI с тёмной темой, прогресс-баром и копируемыми логами.
- Коррекция времени: +4 часа для списка вебинаров, +3 часа для событий.
- Сортировка опросов по времени (до 10 опросов).

## Требования

- **Python**: 3.8 или выше.
- **Библиотеки**:
  - `requests`
  - `openpyxl`
  - `tkinter` (обычно встроен в Python)
- **Zoom API**:
  - Account ID, Client ID, Client Secret, User ID (см. [Zoom API Docs](https://developers.zoom.us/docs/api/)).
- **Операционная система**: Windows, macOS или Linux.

## Установка

1. **Клонируйте репозиторий**:
   ```bash
   git clone https://github.com/your-username/zoom-webinar-report-generator.git
   cd zoom-webinar-report-generator
   ```

2. **Установите зависимости**:
   ```bash
   pip install requests openpyxl
   ```

3. **Настройте Zoom API**:
   - Создайте Server-to-Server OAuth приложение в [Zoom App Marketplace](https://marketplace.zoom.us/).
   - Получите `Account ID`, `Client ID`, `Client Secret` и `User ID`.
   - Сохраните их в `config.ini` (см. ниже) или вводите вручную в GUI.

4. **(Опционально) Создайте `config.ini`**:
   ```ini
   [ZoomCredentials]
   account_id = your_account_id
   client_id = your_client_id
   client_secret = your_client_secret
   user_id = your_user_id
   ```

## Использование

1. **Запустите программу**:
   ```bash
   python zoom_webinar_gui_progress.py
   ```

2. **Интерфейс**:
   - **Поля ввода**:
     - `Account ID`, `Client ID`, `Client Secret`, `User ID` (загружаются из `config.ini`, если есть).
     - `Количество сессий`: 10, 20 или 30.
     - `Директория сохранения`: Выберите папку для отчётов.
   - **Кнопки**:
     - `Загрузить вебинары`: Получает список прошедших вебинаров.
     - `Обработать выбранные вебинары`: Генерирует отчёты для выбранных сессий.
   - **Логи**: Копируйте текст логов через выделение и Ctrl+C.

3. **Процесс**:
   - Заполните поля Zoom API или убедитесь, что `config.ini` настроен.
   - Нажмите `Загрузить вебинары`, выберите вебинары в списке.
   - Нажмите `Обработать выбранные вебинары`.
   - Прогресс отображается в прогресс-баре, логи — в окне.

4. **Выходные файлы**:
   - Папка: `{save_dir}/{YYMMDD}/`
   - Файлы:
     - `Zoom{YYMMDD}.xlsx`: Полный отчёт (15+ столбцов, включая опросы).
     - `roib{YYMMDD}.xlsx`: Упрощённый отчёт (8+ столбцов).
     - `participants_*.json`: Данные участников.
     - `polls_*.json`: Данные опросов.

### Пример выходных данных

- **Zoom250408.xlsx** (фрагмент):
  | Имя пользователя (исходное имя) | Имя   | Фамилия     | Эл. почта          | Роль    | Опрос 1   |
  |--------------------------------|-------|-------------|--------------------|---------|-----------|
  | Pain Russia                    | ROIB  | PainRussia  | pain@example.com   | panelist | 14:30:00  |
  | Иван Иванов                    | Иванов | Иван        | ivan@example.com   | panelist | 14:31:00  |
  | Анна Петрова                   | Анна  | Петрова     | anna@example.com   | attendee | 14:32:00  |

- **roib250408.xlsx** (фрагмент):
  | Время в сеансе (минут) | Фамилия     | Имя   | Дата       | Опрос 1   |
  |------------------------|-------------|-------|------------|-----------|
  | 45.67                  | PainRussia  | ROIB  | 08.04.2025 | 14:30:00  |
  | 30.12                  | Иван        | Иванов | 08.04.2025 | 14:31:00  |
  | 25.89                  | Петрова     | Анна  | 08.04.2025 | 14:32:00  |

## Структура кода

- **zoom_webinar_gui_progress.py**:
  - Класс `ZoomWebinarApp`:
    - `__init__`: Инициализация GUI (тёмная тема, 600x800).
    - `log`: Вывод сообщений в копируемое окно логов.
    - `get_access_token`: Получение токена Zoom API.
    - `get_past_webinars`, `get_webinar_instances`, `get_webinar_participants`, `get_webinar_panelists`, `get_webinar_registrants`, `get_webinar_polls`: Запросы к Zoom API.
    - `merge_participant_data`: Объединение данных с учётом ролей и `"Pain Russia"`.
    - `process_polls`: Обработка опросов (сортировка по `date_time`).
    - `save_to_excel`: Формирование Excel-отчётов.
    - `load_webinars`, `process_selected_webinars`: Управление процессом.

## Ограничения

- Требуется стабильное интернет-соединение для Zoom API.
- Ограниченные данные участников могут повлиять на полноту отчётов (логируются предупреждения).
- Максимум 10 опросов на вебинар.
- Zoom API может иметь лимиты на количество запросов.

## Устранение неполадок

1. **Ошибка "Заполните все поля Zoom API"**:
   - Проверьте `config.ini` или поля ввода.
2. **Ошибка получения токена**:
   - Убедитесь, что `Account ID`, `Client ID`, `Client Secret` корректны.
3. **Пустой список вебинаров**:
   - Проверьте `User ID` и наличие прошедших вебинаров.
4. **Некорректные имена/фамилии**:
   - Проверьте `participants_*.json` на наличие `name`, `first_name`, `last_name`.
5. **Логи не копируются**:
   - Убедитесь, что выделяете текст и используете Ctrl+C.

Для диагностики:
- Скопируйте логи из GUI.
- Проверьте `{save_dir}/{YYMMDD}/participants_*.json`.
- Укажите `webinarId` и `webinarUUID` в issue.

## Лицензия

MIT License. См. [LICENSE](LICENSE) для деталей.

## Контакты

- **Автор**: Гумербаев Айрат
- **GitHub**: [Kelmeer](https://github.com/Kelmeer)
- **Donate**: [Поддержать](https://www.donationalerts.com/r/dr_klmn)

---

⭐ Если программа полезна, поставьте звезду на GitHub!
