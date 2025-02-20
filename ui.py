from PyQt5.QtCore import Qt
from PyQt5.QtWidgets import QApplication, QMainWindow, QVBoxLayout, QPushButton, QLabel, QWidget, QLineEdit, \
    QMessageBox, QHBoxLayout, QTableWidget, QTableWidgetItem, QFileDialog
from matplotlib.backends.backend_qt5agg import FigureCanvasQTAgg as FigureCanvas
from matplotlib.figure import Figure
from psycopg2 import sql
import configparser
from openpyxl import Workbook
import pandas as pd
import auth
import functions
import matplotlib.pyplot as plt


class UserWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setup_menu_ui()
        self.setWindowTitle("Деканат")
        self.resize(600, 400)
        self.center()

    def center(self):
        screen = self.screen().geometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

    def setup_menu_ui(self):
        layout = QVBoxLayout()
        label = QLabel("Добро пожаловать! Выберите действие")
        label.setStyleSheet("font-size: 18px; font-weight: semibold;")
        layout.addWidget(label)

        btn_catalog = QPushButton("Справочники")
        btn_catalog.setStyleSheet("font-size: 16px;")
        btn_catalog.setFixedHeight(40)

        btn_report = QPushButton("Отчеты")
        btn_report.setStyleSheet("font-size: 16px;")
        btn_report.setFixedHeight(40)

        btn_catalog.clicked.connect(self.catalog_ui)
        btn_report.clicked.connect(self.report_ui)

        layout.addWidget(btn_catalog)
        layout.addWidget(btn_report)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def catalog_ui(self):
        layout = QVBoxLayout()

        btn_student_list = QPushButton("Cтуденты")
        btn_student_list.setStyleSheet("font-size: 16px;")
        btn_student_list.setFixedHeight(40)
        btn_teachers_list = QPushButton("Преподаватели")
        btn_teachers_list.setStyleSheet("font-size: 16px;")
        btn_teachers_list.setFixedHeight(40)
        btn_group_list = QPushButton("Группы")
        btn_group_list.setStyleSheet("font-size: 16px;")
        btn_group_list.setFixedHeight(40)
        btn_subjects_list = QPushButton("Предметы")
        btn_subjects_list.setStyleSheet("font-size: 16px;")
        btn_subjects_list.setFixedHeight(40)
        btn_students_marks = QPushButton("Средние оценки студентов")
        btn_students_marks.setStyleSheet("font-size: 16px;")
        btn_students_marks.setFixedHeight(40)
        btn_subjects_marks = QPushButton("Средние оценки по предметам")
        btn_subjects_marks.setStyleSheet("font-size: 16px;")
        btn_subjects_marks.setFixedHeight(40)

        btn_back = QPushButton("Назад")

        btn_student_list.clicked.connect(self.show_students_ui)
        btn_teachers_list.clicked.connect(self.show_teachers_ui)
        btn_group_list.clicked.connect(self.show_group_list_ui)
        btn_subjects_list.clicked.connect(self.show_subjects_list_ui)
        btn_students_marks.clicked.connect(self.students_marks_list_ui)
        btn_subjects_marks.clicked.connect(self.subjects_marks_list_ui)
        btn_back.clicked.connect(self.setup_menu_ui)

        layout.addWidget(btn_student_list)
        layout.addWidget(btn_teachers_list)
        layout.addWidget(btn_group_list)
        layout.addWidget(btn_subjects_list)
        layout.addWidget(btn_students_marks)
        layout.addWidget(btn_subjects_marks)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)


    def show_teachers_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('teachers')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_students_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('students')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_group_list_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('groups')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_subjects_list_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('subjects')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def students_marks_list_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('students_marks')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def subjects_marks_list_ui(self):
        layout = QVBoxLayout()

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        self.load_data_into_table('subjects_marks')

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.catalog_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)


    def report_ui(self):
        layout = QVBoxLayout()

        btn_view_year_grade_avg = QPushButton("""Получить информацию о среднем балле\n в заданном году""")
        btn_view_years_grade_avg = QPushButton("""Получить информацию о средних баллах \n по заданному интервалу""")
        btn_view_student_grade_avg = QPushButton("Получить информацию о студенте и его оценках")
        btn_view_years_subj_avg = QPushButton("""Получить информацию о средних баллах \n по заданному предмету по годам""")
        btn_view_year_grade_avg.setStyleSheet("font-size: 16px;")
        btn_view_years_grade_avg.setStyleSheet("font-size: 16px;")
        btn_view_student_grade_avg.setStyleSheet("font-size: 16px;")
        btn_view_years_subj_avg.setStyleSheet("font-size: 16px;")
        btn_view_year_grade_avg.setFixedHeight(60)
        btn_view_years_grade_avg.setFixedHeight(60)
        btn_view_student_grade_avg.setFixedHeight(60)
        btn_view_years_subj_avg.setFixedHeight(60)

        btn_view_year_grade_avg.clicked.connect(self.show_table_year_grade_ui)
        btn_view_years_grade_avg.clicked.connect(self.show_table_years_grade_ui)
        btn_view_student_grade_avg.clicked.connect(self.show_table_student_grade_ui)
        btn_view_years_subj_avg.clicked.connect(self.show_table_years_subj_ui)

        layout.addWidget(btn_view_year_grade_avg)
        layout.addWidget(btn_view_years_grade_avg)
        layout.addWidget(btn_view_student_grade_avg)
        layout.addWidget(btn_view_years_subj_avg)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_table_years_subj_ui(self):
        layout = QVBoxLayout()

        self.id_input = QLineEdit(self)
        self.id_input.setPlaceholderText("Введите id предмета")
        layout.addWidget(self.id_input)

        btn_add_grade = QPushButton("Получить")
        btn_add_grade.clicked.connect(self.load_years_subject)
        layout.addWidget(btn_add_grade)

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        btn_show_chart = QPushButton("Построить график")
        btn_show_chart.clicked.connect(self.show_chart_years_subject)
        layout.addWidget(btn_show_chart)

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.report_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_table_student_grade_ui(self):
        layout = QVBoxLayout()

        self.id_input = QLineEdit(self)
        self.id_input.setPlaceholderText("Введите id студента")
        layout.addWidget(self.id_input)

        btn_add_grade = QPushButton("Получить")
        btn_add_grade.clicked.connect(self.load_student_grade_into_table)
        layout.addWidget(btn_add_grade)

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        btn_show_chart = QPushButton("Построить график")
        btn_show_chart.clicked.connect(self.show_chart_student_grade)
        layout.addWidget(btn_show_chart)

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.report_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_table_year_grade_ui(self):
        layout = QVBoxLayout()

        self.year_input = QLineEdit(self)
        self.year_input.setPlaceholderText("Введите год")
        layout.addWidget(self.year_input)

        btn_add_grade = QPushButton("Получить")
        btn_add_grade.clicked.connect(self.load_year_into_table)
        layout.addWidget(btn_add_grade)

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        btn_show_chart = QPushButton("Построить график")
        btn_show_chart.clicked.connect(self.show_chart_year)
        layout.addWidget(btn_show_chart)

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.report_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_table_years_grade_ui(self):
        layout = QVBoxLayout()

        self.year1_input = QLineEdit(self)
        self.year1_input.setPlaceholderText("Введите год начала интервала")
        layout.addWidget(self.year1_input)

        self.year2_input = QLineEdit(self)
        self.year2_input.setPlaceholderText("Введите год окончания интервала")
        layout.addWidget(self.year2_input)

        btn_years_grade = QPushButton("Получить")
        btn_years_grade.clicked.connect(self.load_years_into_table)
        layout.addWidget(btn_years_grade)

        self.table_widget = QTableWidget(self)
        layout.addWidget(self.table_widget)

        btn_show_chart = QPushButton("Построить график")
        btn_show_chart.clicked.connect(self.show_chart_years)
        layout.addWidget(btn_show_chart)

        btn_export = QPushButton("Выгрузить в Excel")
        btn_export.clicked.connect(self.export_table_to_excel)
        layout.addWidget(btn_export)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.report_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def load_data_into_table(self, mark):
        if mark == 'teachers':
            rows, columns = functions.load_teachers()
        elif mark == 'students':
            rows, columns = functions.load_students_and_groups()
        elif mark == 'groups':
            rows, columns = functions.load_groups()
        elif mark == 'subjects':
            rows, columns = functions.load_subjects()
        elif mark == 'students_marks':
            rows, columns = functions.load_students_and_marks()
        elif mark == 'subjects_marks':
            rows, columns = functions.load_subjects_and_marks()

        if rows and columns:
            self.table_widget.setRowCount(len(rows))
            self.table_widget.setColumnCount(len(columns))
            self.table_widget.setHorizontalHeaderLabels(columns)

            for row_index, row_data in enumerate(rows):
                for col_index, cell_data in enumerate(row_data):
                    self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(cell_data)))
        else:
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            print("Не удалось загрузить данные в таблицу.")

    def load_years_subject(self):

        id = self.id_input.text()
        rows, columns = functions.load_years_subject(id)

        if rows and columns:
            self.table_widget.setRowCount(len(rows))
            self.table_widget.setColumnCount(len(columns))
            self.table_widget.setHorizontalHeaderLabels(columns)

            for row_index, row_data in enumerate(rows):
                for col_index, cell_data in enumerate(row_data):
                    self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(cell_data)))
        else:
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            print("Не удалось загрузить данные в таблицу.")

    def load_student_grade_into_table(self):

        id = self.id_input.text()
        rows, columns = functions.load_student_grades(id)

        if rows and columns:
            self.table_widget.setRowCount(len(rows))
            self.table_widget.setColumnCount(len(columns))
            self.table_widget.setHorizontalHeaderLabels(columns)

            for row_index, row_data in enumerate(rows):
                for col_index, cell_data in enumerate(row_data):
                    self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(cell_data)))
        else:
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            print("Не удалось загрузить данные в таблицу.")

    def load_year_into_table(self):

        year = self.year_input.text()
        rows, columns = functions.load_year_avg_grade(year)

        if rows and columns:
            self.table_widget.setRowCount(len(rows))
            self.table_widget.setColumnCount(len(columns))
            self.table_widget.setHorizontalHeaderLabels(columns)

            for row_index, row_data in enumerate(rows):
                for col_index, cell_data in enumerate(row_data):
                    self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(cell_data)))
        else:
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            print("Не удалось загрузить данные в таблицу.")

    def load_years_into_table(self):

        year1 = self.year1_input.text()
        year2 = self.year2_input.text()
        rows, columns = functions.load_years_avg_grade(year1, year2)

        if rows and columns:
            self.table_widget.setRowCount(len(rows))
            self.table_widget.setColumnCount(len(columns))
            self.table_widget.setHorizontalHeaderLabels(columns)

            for row_index, row_data in enumerate(rows):
                for col_index, cell_data in enumerate(row_data):
                    self.table_widget.setItem(row_index, col_index, QTableWidgetItem(str(cell_data)))
        else:
            self.table_widget.setRowCount(0)
            self.table_widget.setColumnCount(0)
            print("Не удалось загрузить данные в таблицу.")

    def export_table_to_excel(self):
        file_path, _ = QFileDialog.getSaveFileName(self, "Сохранить файл", "", "Excel Files (*.xlsx)")
        if not file_path:
            return

        row_count = self.table_widget.rowCount()
        column_count = self.table_widget.columnCount()
        data = []

        for row in range(row_count):
            row_data = []
            for column in range(column_count):
                item = self.table_widget.item(row, column)
                row_data.append(item.text() if item else "")
            data.append(row_data)

        headers = [self.table_widget.horizontalHeaderItem(i).text() for i in range(column_count)]

        df = pd.DataFrame(data, columns=headers)

        try:
            df.to_excel(file_path, index=False, engine='openpyxl')
            print(f"Таблица успешно сохранена в файл: {file_path}")
        except Exception as e:
            print(f"Ошибка при сохранении файла: {e}")

    def show_chart_years_subject(self):

        id = self.id_input.text()
        rows, columns = functions.load_years_subject(id)

        if not rows:
            print("Нет данных для построения графика.")
            return

        year = [row[0] for row in rows]
        average_marks = [row[1] for row in rows]

        chart_window = QWidget()
        layout = QVBoxLayout(chart_window)

        figure = Figure(figsize=(10, 6))
        canvas = FigureCanvas(figure)
        layout.addWidget(canvas)

        ax = figure.add_subplot(111)
        ax.plot(year, average_marks, marker='o', linestyle='-', color='b')
        figure.subplots_adjust(bottom=0.3)

        ax.set_title("Средний балл по годам", fontsize=14)
        ax.set_xlabel("Год", fontsize=12)
        ax.set_ylabel("Средний балл", fontsize=12)
        ax.tick_params(axis='x', rotation=45)

        ax.set_xticks(year)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.show_table_years_subj_ui)
        layout.addWidget(btn_back)

        self.setCentralWidget(chart_window)
    def show_chart_student_grade(self):

        id = self.id_input.text()
        rows, columns = functions.load_student_grades(id)

        if not rows:
            print("Нет данных для построения графика.")
            return

        subjects_names = [row[0] for row in rows]
        average_marks = [float(row[1]) for row in rows]

        chart_window = QWidget()
        layout = QVBoxLayout(chart_window)

        figure = Figure(figsize=(10, 6))
        canvas = FigureCanvas(figure)
        layout.addWidget(canvas)

        ax = figure.add_subplot(111)
        ax.plot(subjects_names, average_marks, marker='o', linestyle='-', color='b')
        figure.subplots_adjust(bottom=0.3)

        ax.set_title("Средний балл по предметам", fontsize=14)
        ax.set_xlabel("Предмет", fontsize=12)
        ax.set_ylabel("Средний балл", fontsize=12)
        ax.tick_params(axis='x', rotation=45)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.show_table_student_grade_ui)
        layout.addWidget(btn_back)

        self.setCentralWidget(chart_window)

    def show_chart_year(self):
        year = self.year_input.text()
        rows, columns = functions.load_year_avg_grade(year)

        if not rows:
            print("Нет данных для построения графика.")
            return

        subjects_names = [row[0] for row in rows]
        average_marks = [float(row[1]) for row in rows]

        chart_window = QWidget()
        layout = QVBoxLayout(chart_window)

        figure = Figure(figsize=(10, 6))
        canvas = FigureCanvas(figure)
        layout.addWidget(canvas)

        ax = figure.add_subplot(111)
        ax.plot(subjects_names, average_marks, marker='o', linestyle='-', color='b')
        figure.subplots_adjust(bottom=0.3)

        ax.set_title("Средний балл по предметам", fontsize=14)
        ax.set_xlabel("Предмет", fontsize=12)
        ax.set_ylabel("Средний балл", fontsize=12)
        ax.tick_params(axis='x', rotation=45)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.show_table_year_grade_ui)
        layout.addWidget(btn_back)

        self.setCentralWidget(chart_window)

    def show_chart_years(self):

        year1 = self.year1_input.text()
        year2 = self.year2_input.text()
        rows, columns = functions.load_years_avg_grade(year1, year2)

        if not rows:
            print("Нет данных для построения графика.")
            return

        group_names = [row[0] for row in rows]
        average_marks = [float(row[1]) for row in rows]

        chart_window = QWidget()
        layout = QVBoxLayout(chart_window)

        figure = Figure(figsize=(10, 6))
        canvas = FigureCanvas(figure)
        layout.addWidget(canvas)

        ax = figure.add_subplot(111)
        ax.plot(group_names, average_marks, marker='o', linestyle='-', color='b')
        figure.subplots_adjust(bottom=0.3)

        ax.set_title("Средний балл по предметам", fontsize=14)
        ax.set_xlabel("Группа", fontsize=12)
        ax.set_ylabel("Средний балл", fontsize=12)
        ax.tick_params(axis='x', rotation=45)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.show_table_years_grade_ui)
        layout.addWidget(btn_back)

        self.setCentralWidget(chart_window)

class AdminWindow(UserWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Админ панель")
        self.setup_main_menu_ui()
        self.resize(600, 400)
        self.center()

    def center(self):
        screen = self.screen().geometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)

    def setup_main_menu_ui(self):
        label = QLabel("Добро пожаловать, администратор! Выберите действие")
        label.setStyleSheet("font-size: 18px; font-weight: semibold;")
        layout = QVBoxLayout()
        layout.addWidget(label)

        btn_edit_data = QPushButton("Редактировать данные")
        btn_edit_data.setStyleSheet("font-size: 16px;")
        btn_edit_data.setFixedHeight(40)

        btn_catalog = QPushButton("Справочники")
        btn_catalog.setStyleSheet("font-size: 16px;")
        btn_catalog.setFixedHeight(40)

        btn_report = QPushButton("Отчеты")
        btn_report.setStyleSheet("font-size: 16px;")
        btn_report.setFixedHeight(40)

        btn_edit_data.clicked.connect(self.edit_data_ui)
        btn_catalog.clicked.connect(super().catalog_ui)
        btn_report.clicked.connect(super().report_ui)

        layout.addWidget(btn_edit_data)
        layout.addWidget(btn_catalog)
        layout.addWidget(btn_report)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def edit_data_ui(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Добавление предмета"))

        btn_add_person = QPushButton("Добавить пользователя")
        btn_add_person.setStyleSheet("font-size: 16px;")
        btn_add_person.setFixedHeight(40)
        btn_delete_person = QPushButton("Удалить пользователя")
        btn_delete_person.setStyleSheet("font-size: 16px;")
        btn_delete_person.setFixedHeight(40)
        btn_add_group = QPushButton("Добавить группу")
        btn_add_group.setStyleSheet("font-size: 16px;")
        btn_add_group.setFixedHeight(40)
        btn_delete_group = QPushButton("Удалить группу")
        btn_delete_group.setStyleSheet("font-size: 16px;")
        btn_delete_group.setFixedHeight(40)
        btn_add_subject = QPushButton("Добавить предмет")
        btn_add_subject.setStyleSheet("font-size: 16px;")
        btn_add_subject.setFixedHeight(40)
        btn_delete_subject = QPushButton("Удалить предмет")
        btn_delete_subject.setStyleSheet("font-size: 16px;")
        btn_delete_subject.setFixedHeight(40)
        btn_add_grade = QPushButton("Добавить оценку")
        btn_add_grade.setStyleSheet("font-size: 16px;")
        btn_add_grade.setFixedHeight(40)
        btn_delete_grade = QPushButton("Удалить оценку")
        btn_delete_grade.setStyleSheet("font-size: 16px;")
        btn_delete_grade.setFixedHeight(40)

        btn_add_person.clicked.connect(self.show_add_person_ui)
        btn_delete_person.clicked.connect(self.show_delete_person_ui)
        btn_add_group.clicked.connect(self.show_add_group_ui)
        btn_delete_group.clicked.connect(self.show_delete_group_ui)
        btn_add_subject.clicked.connect(self.show_add_subject_ui)
        btn_delete_subject.clicked.connect(self.show_delete_subject_ui)
        btn_add_grade.clicked.connect(self.show_add_grade_ui)
        btn_delete_grade.clicked.connect(self.show_delete_grade_ui)

        h_layout1 = QHBoxLayout()
        h_layout1.addWidget(btn_add_person)
        h_layout1.addWidget(btn_delete_person)

        h_layout2 = QHBoxLayout()
        h_layout2.addWidget(btn_add_group)
        h_layout2.addWidget(btn_delete_group)

        h_layout3 = QHBoxLayout()
        h_layout3.addWidget(btn_add_subject)
        h_layout3.addWidget(btn_delete_subject)

        h_layout4 = QHBoxLayout()
        h_layout4.addWidget(btn_add_grade)
        h_layout4.addWidget(btn_delete_grade)

        layout.addLayout(h_layout1)
        layout.addLayout(h_layout2)
        layout.addLayout(h_layout3)
        layout.addLayout(h_layout4)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_add_subject_ui(self):
        layout = QVBoxLayout()

        self.subject_name_input = QLineEdit(self)
        self.subject_name_input.setPlaceholderText("Введите название предмета")
        layout.addWidget(self.subject_name_input)

        btn_add_subject = QPushButton("Добавить")
        btn_add_subject.clicked.connect(self.add_subject)
        layout.addWidget(btn_add_subject)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_delete_subject_ui(self):
        layout = QVBoxLayout()

        self.subject_id_input = QLineEdit(self)
        self.subject_id_input.setPlaceholderText("Введите id предмета")
        layout.addWidget(self.subject_id_input)

        btn_delete_subject = QPushButton("Добавить")
        btn_delete_subject.clicked.connect(self.delete_subject)
        layout.addWidget(btn_delete_subject)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_add_grade_ui(self):
        layout = QVBoxLayout()

        self.student_id_input = QLineEdit(self)
        self.student_id_input.setPlaceholderText("Введите id студента")
        layout.addWidget(self.student_id_input)

        self.subject_id_input = QLineEdit(self)
        self.subject_id_input.setPlaceholderText("Введите id предмета")
        layout.addWidget(self.subject_id_input)

        self.teacher_id_input = QLineEdit(self)
        self.teacher_id_input.setPlaceholderText("Введите id преподавателя")
        layout.addWidget(self.teacher_id_input)

        self.grade_value_input = QLineEdit(self)
        self.grade_value_input.setPlaceholderText("Введите оценку")
        layout.addWidget(self.grade_value_input)

        btn_add_grade = QPushButton("Добавить")
        btn_add_grade.clicked.connect(self.add_grade)
        layout.addWidget(btn_add_grade)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_delete_grade_ui(self):
        layout = QVBoxLayout()

        self.grade_id_input = QLineEdit(self)
        self.grade_id_input.setPlaceholderText("Введите id оценки")
        layout.addWidget(self.grade_id_input)

        btn_delete_grade = QPushButton("Удалить")
        btn_delete_grade.clicked.connect(self.delete_grade)
        layout.addWidget(btn_delete_grade)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)
    def show_add_group_ui(self):
        layout = QVBoxLayout()

        self.group_name_input = QLineEdit(self)
        self.group_name_input.setPlaceholderText("Введите номер группы")
        layout.addWidget(self.group_name_input)

        btn_add_group = QPushButton("Добавить")
        btn_add_group.clicked.connect(self.add_group)
        layout.addWidget(btn_add_group)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_delete_group_ui(self):
        layout = QVBoxLayout()

        self.group_id_input = QLineEdit(self)
        self.group_id_input.setPlaceholderText("Введите id группы")
        layout.addWidget(self.group_id_input)

        btn_delete_group = QPushButton("Удалить")
        btn_delete_group.clicked.connect(self.delete_group)
        layout.addWidget(btn_delete_group)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_delete_person_ui(self):
        layout = QVBoxLayout()

        self.person_id_input = QLineEdit(self)
        self.person_id_input.setPlaceholderText("Введите id пользователя")
        layout.addWidget(self.person_id_input)

        btn_delete_person = QPushButton("Удалить")
        btn_delete_person.clicked.connect(self.delete_person)
        layout.addWidget(btn_delete_person)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def show_add_person_ui(self):
        layout = QVBoxLayout()
        layout.addWidget(QLabel("Добавление нового студента/преподавателя"))

        self.person_name_input = QLineEdit(self)
        self.person_name_input.setPlaceholderText("Введите имя")
        layout.addWidget(self.person_name_input)

        self.person_last_name_input = QLineEdit(self)
        self.person_last_name_input.setPlaceholderText("Введите фамилию")
        layout.addWidget(self.person_last_name_input)

        self.person_father_name_input = QLineEdit(self)
        self.person_father_name_input.setPlaceholderText("Введите отчество")
        layout.addWidget(self.person_father_name_input)

        self.person_group_id_input = QLineEdit(self)
        self.person_group_id_input.setPlaceholderText("Введите ID группы")
        layout.addWidget(self.person_group_id_input)

        self.person_type_input = QLineEdit(self)
        self.person_type_input.setPlaceholderText("Введите тип")
        layout.addWidget(self.person_type_input)

        btn_add_person = QPushButton("Добавить")
        btn_add_person.clicked.connect(self.add_person)
        layout.addWidget(btn_add_person)

        btn_back = QPushButton("Назад")
        btn_back.clicked.connect(self.setup_main_menu_ui)
        layout.addWidget(btn_back)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def add_grade(self):
        student_id = self.student_id_input.text()
        subject_id = self.subject_id_input.text()
        teacher_id = self.student_id_input.text()
        grade_value = self.grade_value_input.text()

        if student_id and subject_id and teacher_id and grade_value:
            result = functions.add_grade(student_id, subject_id, teacher_id, grade_value)
            if result:
                QMessageBox.information(self, "Успех", "Оценка добавлена!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить оценку.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")
    def delete_grade(self):
        grade_id = self.grade_id_input.text()

        if grade_id:
            result = functions.delete_grade(grade_id)
            if result:
                QMessageBox.information(self, "Успех", "Оценка удалена!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить оценку.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")
    def add_subject(self):
        subject_name = self.subject_name_input.text()

        if subject_name:
            result = functions.add_subject(subject_name)
            if result:
                QMessageBox.information(self, "Успех", "Предмет добавлен!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить предмет.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")

    def delete_subject(self):
        subject_id = self.subject_id_input.text()

        if subject_id:
            result = functions.delete_subject(subject_id)
            if result:
                QMessageBox.information(self, "Успех", "Предмет удален!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить предмет.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")

    def add_group(self):
        group_name = self.group_name_input.text()

        if group_name:
            result = functions.add_group(group_name)
            if result:
                QMessageBox.information(self, "Успех", "Группа добавлена!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить группу.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")
    def add_person(self):
        last_name = self.person_last_name_input.text()
        first_name = self.person_name_input.text()
        father_name = self.person_father_name_input.text()
        group_id = self.person_group_id_input.text()
        type = self.person_type_input.text()

        if last_name and first_name and father_name and group_id and type:
            result = functions.add_person(last_name, first_name, father_name, group_id, type)
            if result:
                QMessageBox.information(self, "Успех", "Студент добавлен!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось добавить студента.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите все данные!")

    def delete_person(self):
        person_id = self.person_id_input.text()
        if person_id:
            result = functions.delete_person(person_id)
            if result:
                QMessageBox.information(self, "Успех", "Пользователь удален!")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить пользователя.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")

    def delete_group(self):
        group_id = self.group_id_input.text()
        if group_id:
            result = functions.delete_group(group_id)
            if result:
                QMessageBox.information(self, "Успех", "Группа удалена")
            else:
                QMessageBox.warning(self, "Ошибка", "Не удалось удалить группу.")
        else:
            QMessageBox.warning(self, "Ошибка", "Пожалуйста, введите данные!")


class LoginWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Авторизация")
        self.setup_login_ui()
        self.resize(400, 300)
        self.center()

    def center(self):
        screen = self.screen().geometry()
        window = self.geometry()
        x = (screen.width() - window.width()) // 2
        y = (screen.height() - window.height()) // 2
        self.move(x, y)
    def setup_login_ui(self):
        layout = QVBoxLayout()
        layout.setAlignment(Qt.AlignCenter)

        self.username_input = QLineEdit(self)
        self.username_input.setPlaceholderText("Введите логин")
        self.username_input.setFixedWidth(200)
        self.username_input.setStyleSheet("font-size: 16px; padding: 5px;")
        layout.addWidget(self.username_input)

        self.password_input = QLineEdit(self)
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setPlaceholderText("Введите пароль")
        self.password_input.setFixedWidth(200)
        self.password_input.setStyleSheet("font-size: 16px; padding: 5px;")
        layout.addWidget(self.password_input)

        btn_login = QPushButton("Войти")
        btn_login.setFixedSize(100, 40)
        btn_login.clicked.connect(self.login)
        btn_login.setStyleSheet("font-size: 16px;")
        layout.addWidget(btn_login, alignment=Qt.AlignCenter)

        container = QWidget()
        container.setLayout(layout)
        self.setCentralWidget(container)

    def login(self):
        username = self.username_input.text()
        password = self.password_input.text()

        role = auth.authenticate_user(username, password)

        if role == 'admin':
            self.window = AdminWindow()
            self.window.show()
            self.close()
        elif role == 'user':
            self.window = UserWindow()
            self.window.show()
            self.close()
        else:
            QMessageBox.warning(self, "Ошибка", "Неверный логин или пароль!")