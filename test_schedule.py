import unittest
import os
import json
import shutil
from collections import defaultdict
from ortools.sat.python import cp_model # Залишимо імпорт для повної сумісності, хоча в _create_lecture_objects_for_test він не використовується напряму

# --- Перевизначення необхідних частин з основного скрипту для тестування ---
# В реальному проекті ці функції імпортувались би з окремого модуля (наприклад, schedule_app.py).
# Для демонстрації в межах Canvas, ми включаємо їх тут або створюємо спрощені версії.

# Константи та клас Lecture
DAYS = ["Пн", "Вт", "Ср", "Чт", "Пт"]
# SLOTS_PER_DAY є динамічним, тому ми передаємо його до функції напряму.

class Lecture:
    """Представляє одну лекцію (пару) з усіма її атрибутами."""
    def __init__(self, group, subject, teacher, count):
        self.group = group
        self.subject = subject
        self.teacher = teacher
        self.count = count # Кількість годин/пар на тиждень для цього предмета
        self.vars = [] # Змінні CP-SAT для слотів і кімнат цієї лекції (заповнюються пізніше)

def load_json_for_test(path):
    """
    Допоміжна функція для завантаження JSON-файлів під час тестування.
    Не використовує messagebox для уникнення GUI під час тестів.
    """
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except FileNotFoundError:
        raise FileNotFoundError(f"Тестовий файл не знайдено: {path}")
    except json.JSONDecodeError:
        raise json.JSONDecodeError(f"Помилка декодування JSON у тестовому файлі: {path}", doc=path, pos=0)


# Допоміжна функція, що імітує логіку створення об'єктів Lecture з run_solver_and_generate_reports
def _create_lecture_objects_for_test(data_folder, mock_slots_per_day):
    """
    Створює об'єкти Lecture на основі тестових JSON-даних.
    Це ізольована частина логіки з run_solver_and_generate_reports,
    що стосується лише створення об'єктів Lecture та базових перевірок.
    """
    groups = load_json_for_test(os.path.join(data_folder, "groups.json"))
    subjects = load_json_for_test(os.path.join(data_folder, "subjects.json")) # subjects потрібні для subject_types, але не для створення Lecture безпосередньо
    teachers = load_json_for_test(os.path.join(data_folder, "teachers.json")) # teachers потрібні для перевірки існування

    lectures = []
    _total_slots = len(DAYS) * mock_slots_per_day
    
    # Створення об'єктів Lecture на основі вхідних даних
    for group in groups:
        for subj in group["subjects"]:
            teacher = subj["teacher"]
            name = subj["name"]
            count = subj["hours"]
            # Перевірка на перевантаження годин, аналогічно оригінальній функції
            if count > _total_slots:
                raise ValueError(f"Предмет '{name}' для групи '{group['name']}' має {count} годин, що перевищує загальну доступну кількість слотів ({_total_slots}).")
            lectures.append(Lecture(group["name"], name, teacher, count))
    return lectures


class TestLectureCreation(unittest.TestCase):

    def setUp(self):
        """
        Налаштування тестового середовища:
        Створення тимчасової папки та тестових JSON-файлів.
        """
        self.test_data_dir = "test_data_for_lecture"
        os.makedirs(self.test_data_dir, exist_ok=True)

        self.groups_data = [
            {"name": "Група_Тест_А", "subjects": [{"name": "Математика_Т", "teacher": "Петров_Т", "hours": 3}]},
            {"name": "Група_Тест_Б", "subjects": [{"name": "Фізика_Т", "teacher": "Сидоров_Т", "hours": 2}]}
        ]
        self.teachers_data = [{"name": "Петров_Т"}, {"name": "Сидоров_Т"}]
        self.subjects_data = [{"name": "Математика_Т", "type": "лекція"}, {"name": "Фізика_Т", "type": "практика"}]
        self.rooms_data = [{"name": "Ауд_Т1", "type": "лекція"}, {"name": "Лаб_Т2", "type": "практика"}]

        with open(os.path.join(self.test_data_dir, "groups.json"), "w", encoding="utf-8") as f:
            json.dump(self.groups_data, f, ensure_ascii=False, indent=4)
        with open(os.path.join(self.test_data_dir, "teachers.json"), "w", encoding="utf-8") as f:
            json.dump(self.teachers_data, f, ensure_ascii=False, indent=4)
        with open(os.path.join(self.test_data_dir, "subjects.json"), "w", encoding="utf-8") as f:
            json.dump(self.subjects_data, f, ensure_ascii=False, indent=4)
        with open(os.path.join(self.test_data_dir, "rooms.json"), "w", encoding="utf-8") as f:
            json.dump(self.rooms_data, f, ensure_ascii=False, indent=4)

    def tearDown(self):
        """
        Очищення тестового середовища:
        Видалення тимчасової папки та її вмісту.
        """
        shutil.rmtree(self.test_data_dir) # Використовуємо shutil.rmtree для видалення директорії та її вмісту

    def test_lecture_objects_are_created_correctly(self):
        """
        Перевіряє, чи коректно створюються об'єкти Lecture на основі тестових даних.
        """
        mock_slots_per_day = 5
        lectures = _create_lecture_objects_for_test(self.test_data_dir, mock_slots_per_day)

        # Перевірка загальної кількості створених об'єктів Lecture
        # Очікуємо 2 лекції: 1 для Групи_Тест_А (Математика_Т) і 1 для Групи_Тест_Б (Фізика_Т)
        self.assertEqual(len(lectures), 2, "Має бути створено 2 об'єкти Lecture.")

        # Перевірка атрибутів конкретної лекції для Групи_Тест_А та Математики_Т
        found_math = False
        for lec in lectures:
            if lec.group == "Група_Тест_А" and lec.subject == "Математика_Т":
                self.assertEqual(lec.teacher, "Петров_Т", "Неправильний викладач для Математики_Т.")
                self.assertEqual(lec.count, 3, "Неправильна кількість годин для Математики_Т.")
                self.assertEqual(len(lec.vars), 0, "Змінні vars повинні бути пустими після ініціалізації Lecture.")
                found_math = True
                break
        self.assertTrue(found_math, "Об'єкт Lecture для Математики_Т не знайдено.")

        # Перевірка атрибутів конкретної лекції для Групи_Тест_Б та Фізики_Т
        found_physics = False
        for lec in lectures:
            if lec.group == "Група_Тест_Б" and lec.subject == "Фізика_Т":
                self.assertEqual(lec.teacher, "Сидоров_Т", "Неправильний викладач для Фізики_Т.")
                self.assertEqual(lec.count, 2, "Неправильна кількість годин для Фізики_Т.")
                self.assertEqual(len(lec.vars), 0, "Змінні vars повинні бути пустими після ініціалізації Lecture.")
                found_physics = True
                break
        self.assertTrue(found_physics, "Об'єкт Lecture для Фізики_Т не знайдено.")

    def test_lecture_creation_with_excessive_hours(self):
        """
        Тестує випадок, коли задана кількість годин перевищує загальну доступну
        кількість слотів, що повинно викликати ValueError.
        """
        # Створення тестових даних з надмірною кількістю годин
        excessive_hours_data = [
            {"name": "Група_Перевантажена", "subjects": [{"name": "Забагато_Годин", "teacher": "Тест_Вчитель", "hours": 100}]}
        ]
        with open(os.path.join(self.test_data_dir, "groups.json"), "w", encoding="utf-8") as f:
            json.dump(excessive_hours_data, f, ensure_ascii=False, indent=4)
        
        # Очікуємо, що буде викликано ValueError
        with self.assertRaises(ValueError) as cm:
            _create_lecture_objects_for_test(self.test_data_dir, 5) # 5 слотів/день * 5 днів = 25 загальних слотів, 100 годин забагато
        
        # Перевіряємо, що повідомлення про помилку містить очікуваний текст
        self.assertIn("перевищує загальну доступну кількість слотів", str(cm.exception))

    def test_lecture_creation_with_missing_files(self):
        """
        Тестує, що відсутність JSON-файлів викликає FileNotFoundError.
        (Цей тест перевіряє _create_lecture_objects_for_test,
        а не повний run_solver_and_generate_reports, де є messagebox.)
        """
        # Видаляємо один з файлів, щоб імітувати його відсутність
        os.remove(os.path.join(self.test_data_dir, "groups.json"))
        
        # Очікуємо, що буде викликано FileNotFoundError
        with self.assertRaises(FileNotFoundError) as cm:
            _create_lecture_objects_for_test(self.test_data_dir, 5)
        
        self.assertIn("Тестовий файл не знайдено", str(cm.exception))


if __name__ == '__main__':
    unittest.main()
