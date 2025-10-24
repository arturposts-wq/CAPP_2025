import sys
import os
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, 
                             QLabel, QComboBox, QPushButton, QTableWidget, QTableWidgetItem, 
                             QFileDialog, QMessageBox, QDialog, QFormLayout, QTabWidget, 
                             QInputDialog, QLineEdit, QSpinBox, QSpacerItem, QSizePolicy, 
                             QGroupBox, QScrollArea, QFrame)
from PyQt5.QtCore import Qt
import sqlite3
from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors
from datetime import datetime
import openpyxl
import traceback

class CAPPDatabase:
    def __init__(self, db_name="capp.db"):
        try:
            self.conn = sqlite3.connect(db_name)
            self.cursor = self.conn.cursor()
            self.create_tables()
            print("База данных успешно инициализирована.")
        except sqlite3.Error as e:
            print(f"Ошибка базы данных: {e}")
            raise

    def create_tables(self):
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS models (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL UNIQUE
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS parts (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model_id INTEGER,
                name TEXT NOT NULL,
                code TEXT,
                quantity INTEGER,
                FOREIGN KEY (model_id) REFERENCES models(id)
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS operations (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model_id INTEGER,
                number TEXT,
                code TEXT,
                name TEXT NOT NULL,
                description TEXT,
                equipment TEXT,
                document TEXT,
                prep_time REAL,
                unit_time REAL,
                FOREIGN KEY (model_id) REFERENCES models(id)
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS workshop (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                workshop_name TEXT NOT NULL,
                section TEXT,
                rm TEXT
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS equipment (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                article TEXT,
                note TEXT
            )
        ''')
        self.cursor.execute('''
            CREATE TABLE IF NOT EXISTS document_details (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                model_id INTEGER,
                organization TEXT NOT NULL,
                product_code TEXT,
                document_code TEXT,
                developed_by TEXT,
                checked_by TEXT,
                FOREIGN KEY (model_id) REFERENCES models(id)
            )
        ''')
        self.conn.commit()

    def insert_model(self, name):
        try:
            self.cursor.execute("INSERT OR IGNORE INTO models (name) VALUES (?)", (name,))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки модели: {e}")
            return None

    def insert_part(self, model_id, name, code, quantity):
        try:
            self.cursor.execute("INSERT INTO parts (model_id, name, code, quantity) VALUES (?, ?, ?, ?)",
                               (model_id, name, code, quantity))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки детали: {e}")
            return None

    def insert_operation(self, model_id, number, code, name, description):
        try:
            self.cursor.execute("INSERT INTO operations (model_id, number, code, name, description, equipment, document, prep_time, unit_time) VALUES (?, ?, ?, ?, ?, NULL, NULL, 0.0, 0.0)",
                               (model_id, number, code, name, description))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки операции: {e}")
            return None

    def insert_workshop(self, workshop_name, section, rm):
        try:
            self.cursor.execute("INSERT INTO workshop (workshop_name, section, rm) VALUES (?, ?, ?)",
                               (workshop_name, section, rm))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки данных расцеховки: {e}")
            return None

    def insert_equipment(self, name, article, note):
        try:
            self.cursor.execute("INSERT INTO equipment (name, article, note) VALUES (?, ?, ?)",
                               (name, article, note))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки оборудования: {e}")
            return None

    def insert_document_details(self, model_id, organization, product_code, document_code, developed_by, checked_by):
        try:
            self.cursor.execute("INSERT INTO document_details (model_id, organization, product_code, document_code, developed_by, checked_by) VALUES (?, ?, ?, ?, ?, ?)",
                               (model_id, organization, product_code, document_code, developed_by, checked_by))
            self.conn.commit()
            return self.cursor.lastrowid
        except sqlite3.Error as e:
            print(f"Ошибка вставки реквизитов документа: {e}")
            return None

    def update_model(self, id, name):
        try:
            self.cursor.execute("UPDATE models SET name = ? WHERE id = ?", (name, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления модели: {e}")
            return False

    def update_part(self, id, name, code, quantity):
        try:
            self.cursor.execute("UPDATE parts SET name = ?, code = ?, quantity = ? WHERE id = ?", (name, code, quantity, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления детали: {e}")
            return False

    def update_operation(self, id, number, code, name, description):
        try:
            self.cursor.execute("UPDATE operations SET number = ?, code = ?, name = ?, description = ? WHERE id = ?",
                               (number, code, name, description, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления операции: {e}")
            return False

    def update_workshop(self, id, workshop_name, section, rm):
        try:
            self.cursor.execute("UPDATE workshop SET workshop_name = ?, section = ?, rm = ? WHERE id = ?", 
                               (workshop_name, section, rm, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления данных расцеховки: {e}")
            return False

    def update_equipment(self, id, name, article, note):
        try:
            self.cursor.execute("UPDATE equipment SET name = ?, article = ?, note = ? WHERE id = ?", 
                               (name, article, note, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления оборудования: {e}")
            return False

    def update_document_details(self, id, organization, product_code, document_code, developed_by, checked_by):
        try:
            self.cursor.execute("UPDATE document_details SET organization = ?, product_code = ?, document_code = ?, developed_by = ?, checked_by = ? WHERE id = ?",
                               (organization, product_code, document_code, developed_by, checked_by, id))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка обновления реквизитов документа: {e}")
            return False

    def delete_model(self, id):
        try:
            self.cursor.execute("DELETE FROM parts WHERE model_id = ?", (id,))
            self.cursor.execute("DELETE FROM operations WHERE model_id = ?", (id,))
            self.cursor.execute("DELETE FROM document_details WHERE model_id = ?", (id,))
            self.cursor.execute("DELETE FROM models WHERE id = ?", (id,))
            self.conn.commit()
            return True
        except sqlite3.Error as e:
            print(f"Ошибка удаления модели: {e}")
            return False

    def delete_part(self, id):
        try:
            self.cursor.execute("DELETE FROM parts WHERE id = ?", (id,))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка удаления детали: {e}")
            return False

    def delete_operation(self, id):
        try:
            self.cursor.execute("DELETE FROM operations WHERE id = ?", (id,))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка удаления операции: {e}")
            return False

    def delete_workshop(self, id):
        try:
            self.cursor.execute("DELETE FROM workshop WHERE id = ?", (id,))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка удаления данных расцеховки: {e}")
            return False

    def delete_equipment(self, id):
        try:
            self.cursor.execute("DELETE FROM equipment WHERE id = ?", (id,))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка удаления оборудования: {e}")
            return False

    def delete_document_details(self, id):
        try:
            self.cursor.execute("DELETE FROM document_details WHERE id = ?", (id,))
            self.conn.commit()
            return self.cursor.rowcount > 0
        except sqlite3.Error as e:
            print(f"Ошибка удаления реквизитов документа: {e}")
            return False

    def get_models(self):
        try:
            self.cursor.execute("SELECT id, name FROM models")
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения моделей: {e}")
            return []

    def get_model_id(self, name):
        try:
            self.cursor.execute("SELECT id FROM models WHERE name = ?", (name,))
            result = self.cursor.fetchone()
            return result[0] if result else None
        except sqlite3.Error as e:
            print(f"Ошибка получения ID модели: {e}")
            return None

    def get_parts(self, model_id):
        try:
            self.cursor.execute("SELECT id, name, code, quantity FROM parts WHERE model_id = ?", (model_id,))
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения деталей: {e}")
            return []

    def get_operations(self, model_id=None):
        try:
            if model_id is None:
                self.cursor.execute("SELECT id, number, code, name, description, equipment, document, prep_time, unit_time FROM operations WHERE model_id IS NULL")
                return self.cursor.fetchall()
            else:
                self.cursor.execute("SELECT id, number, code, name, description, equipment, document, prep_time, unit_time FROM operations WHERE model_id = ?", (model_id,))
                return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения операций: {e}")
            return []

    def get_workshop(self):
        try:
            self.cursor.execute("SELECT id, workshop_name, section, rm FROM workshop")
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения данных расцеховки: {e}")
            return []

    def get_equipment(self):
        try:
            self.cursor.execute("SELECT id, name, article, note FROM equipment")
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения оборудования: {e}")
            return []

    def get_document_details(self, model_id):
        try:
            self.cursor.execute("SELECT id, organization, product_code, document_code, developed_by, checked_by FROM document_details WHERE model_id = ?", (model_id,))
            return self.cursor.fetchall()
        except sqlite3.Error as e:
            print(f"Ошибка получения реквизитов документа: {e}")
            return []

    def import_from_excel(self, file_path, current_model=None):
        try:
            wb = openpyxl.load_workbook(file_path)
            imported = False
            if 'Лист1' in wb.sheetnames:
                ws = wb['Лист1']
                headers = [cell.value for cell in ws[1] if cell.value]
                required = ['№', 'Номенклатура', 'Количество']
                if all(col in headers for col in required):
                    for row in ws.iter_rows(min_row=2, values_only=True):
                        code, name, quantity = [row[headers.index(col)] for col in required]
                        if name and quantity is not None:
                            if current_model:
                                model_id = self.get_model_id(current_model)
                                if model_id:
                                    self.insert_part(model_id, name, code, quantity)
                                    imported = True
                else:
                    print("Ошибка: В листе 'Лист1' отсутствуют колонки: №, Номенклатура, Количество")
            else:
                print("Ошибка: Лист 'Лист1' не найден")
            return imported
        except Exception as e:
            print(f"Ошибка импорта из Excel: {e}")
            return False

    def close(self):
        self.conn.close()

class OperationDialog(QDialog):
    def __init__(self, parent=None, is_edit_db=False, db=None):
        super().__init__(parent)
        self.is_edit_db = is_edit_db
        self.db = db
        self.setWindowTitle("Добавить операцию")
        self.setGeometry(200, 200, 400, 150 if is_edit_db else 200)
        layout = QFormLayout()

        if not is_edit_db:
            self.number_input = QLineEdit(self)
            layout.addRow("Номер:", self.number_input)

        self.code_combo = QComboBox(self)
        self.code_combo.setEditable(True)
        self.code_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addRow("Код:", self.code_combo)

        self.name_combo = QComboBox(self)
        self.name_combo.setEditable(True)
        self.name_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addRow("Наименование:", self.name_combo)

        if not is_edit_db:
            self.description_input = QLineEdit(self)
            layout.addRow("Описание:", self.description_input)

        buttons = QHBoxLayout()
        ok_button = QPushButton("ОК", self)
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Отмена", self)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)
        main_layout.addLayout(buttons)
        self.setLayout(main_layout)

        if db:
            self.populate_combos()
            self.code_combo.currentTextChanged.connect(self.sync_name)
            self.name_combo.currentTextChanged.connect(self.sync_code)

    def populate_combos(self):
        operations = self.db.get_operations()
        self.code_combo.clear()
        self.name_combo.clear()
        self.operation_map = {}
        for op in operations:
            code = op[2] or ""
            name = op[3] or ""
            self.operation_map[code] = name
            self.operation_map[name] = code
            if code and self.code_combo.findText(code) == -1:
                self.code_combo.addItem(code)
            if name and self.name_combo.findText(name) == -1:
                self.name_combo.addItem(name)

    def sync_name(self, code):
        if code in self.operation_map:
            name = self.operation_map[code]
            if name and self.name_combo.currentText() != name:
                self.name_combo.setCurrentText(name)

    def sync_code(self, name):
        if name in self.operation_map:
            code = self.operation_map[name]
            if code and self.code_combo.currentText() != code:
                self.code_combo.setCurrentText(code)

    def get_values(self):
        if self.is_edit_db:
            return (self.code_combo.currentText(), self.name_combo.currentText())
        else:
            return (self.number_input.text(), self.code_combo.currentText(), self.name_combo.currentText(), self.description_input.text())

class WorkshopDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить расцеховку")
        self.setGeometry(200, 200, 300, 150)
        layout = QFormLayout()

        self.workshop_input = QLineEdit(self)
        layout.addRow("Цех:", self.workshop_input)

        self.section_input = QLineEdit(self)
        layout.addRow("Участок:", self.section_input)

        self.rm_input = QLineEdit(self)
        layout.addRow("РМ:", self.rm_input)

        buttons = QHBoxLayout()
        ok_button = QPushButton("ОК", self)
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Отмена", self)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)
        main_layout.addLayout(buttons)
        self.setLayout(main_layout)

    def get_values(self):
        return self.workshop_input.text(), self.section_input.text(), self.rm_input.text()

class EquipmentDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить оборудование")
        self.setGeometry(200, 200, 300, 150)
        layout = QFormLayout()

        self.name_input = QLineEdit(self)
        layout.addRow("Наименование:", self.name_input)

        self.article_input = QLineEdit(self)
        layout.addRow("Артикул:", self.article_input)

        self.note_input = QLineEdit(self)
        layout.addRow("Примечание:", self.note_input)

        buttons = QHBoxLayout()
        ok_button = QPushButton("ОК", self)
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Отмена", self)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)
        main_layout.addLayout(buttons)
        self.setLayout(main_layout)

    def get_values(self):
        return self.name_input.text(), self.article_input.text(), self.note_input.text()

class DocumentDetailsDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Добавить реквизиты документа")
        self.setGeometry(200, 200, 400, 200)
        layout = QFormLayout()

        self.organization_input = QLineEdit(self)
        layout.addRow("Организация:", self.organization_input)

        self.product_code_input = QLineEdit(self)
        layout.addRow("Обозначение изделия:", self.product_code_input)

        self.document_code_input = QLineEdit(self)
        layout.addRow("Обозначение документа:", self.document_code_input)

        self.developed_by_input = QLineEdit(self)
        layout.addRow("Разработал:", self.developed_by_input)

        self.checked_by_input = QLineEdit(self)
        layout.addRow("Проверил:", self.checked_by_input)

        buttons = QHBoxLayout()
        ok_button = QPushButton("ОК", self)
        ok_button.clicked.connect(self.accept)
        cancel_button = QPushButton("Отмена", self)
        cancel_button.clicked.connect(self.reject)
        buttons.addWidget(ok_button)
        buttons.addWidget(cancel_button)

        main_layout = QVBoxLayout()
        main_layout.addLayout(layout)
        main_layout.addLayout(buttons)
        self.setLayout(main_layout)

    def get_values(self):
        return (self.organization_input.text(), self.product_code_input.text(), 
                self.document_code_input.text(), self.developed_by_input.text(), 
                self.checked_by_input.text())

class EditDBDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Редактировать БД")
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()

        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        self.operations_tab = QWidget()
        self.tab_widget.addTab(self.operations_tab, "Операции")
        self.setup_operations_tab()

        self.workshop_tab = QWidget()
        self.tab_widget.addTab(self.workshop_tab, "Расцеховка")
        self.setup_workshop_tab()

        self.equipment_tab = QWidget()
        self.tab_widget.addTab(self.equipment_tab, "Оборудование")
        self.setup_equipment_tab()

        buttons = QHBoxLayout()
        close_button = QPushButton("Закрыть")
        close_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #FF9800; color: white; border-radius: 5px;")
        close_button.clicked.connect(self.accept)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self.setLayout(layout)

    def setup_operations_tab(self):
        layout = QVBoxLayout()
        self.operations_table = QTableWidget()
        self.operations_table.setColumnCount(2)
        self.operations_table.setHorizontalHeaderLabels(['Код', 'Наименование'])
        self.operations_table.horizontalHeader().setStretchLastSection(True)
        self.operations_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.operations_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_operation)
        buttons.addWidget(add_button)
        edit_operation_button = QPushButton("Редактировать")
        edit_operation_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_operation_button.clicked.connect(self.edit_operation)
        buttons.addWidget(edit_operation_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_operation)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.operations_tab.setLayout(layout)
        self.update_operations_table()

    def setup_workshop_tab(self):
        layout = QVBoxLayout()
        self.workshop_table = QTableWidget()
        self.workshop_table.setColumnCount(3)
        self.workshop_table.setHorizontalHeaderLabels(['Цех', 'Участок', 'РМ'])
        self.workshop_table.horizontalHeader().setStretchLastSection(True)
        self.workshop_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.workshop_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_workshop)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_workshop)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_workshop)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.workshop_tab.setLayout(layout)
        self.update_workshop_table()

    def setup_equipment_tab(self):
        layout = QVBoxLayout()
        self.equipment_table = QTableWidget()
        self.equipment_table.setColumnCount(3)
        self.equipment_table.setHorizontalHeaderLabels(['Наименование', 'Артикул', 'Примечание'])
        self.equipment_table.horizontalHeader().setStretchLastSection(True)
        self.equipment_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.equipment_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_equipment)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_equipment)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_equipment)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.equipment_tab.setLayout(layout)
        self.update_equipment_table()

    def update_operations_table(self):
        operations = self.db.get_operations()
        self.operations_table.setRowCount(len(operations))
        for row, (id, _, code, name, _, _, _, _, _) in enumerate(operations):
            self.operations_table.setItem(row, 0, QTableWidgetItem(code or ""))
            self.operations_table.setItem(row, 1, QTableWidgetItem(name))
            self.operations_table.item(row, 0).setData(Qt.UserRole, id)

    def update_workshop_table(self):
        workshops = self.db.get_workshop()
        self.workshop_table.setRowCount(len(workshops))
        for row, (id, workshop_name, section, rm) in enumerate(workshops):
            self.workshop_table.setItem(row, 0, QTableWidgetItem(workshop_name))
            self.workshop_table.setItem(row, 1, QTableWidgetItem(section or ""))
            self.workshop_table.setItem(row, 2, QTableWidgetItem(rm or ""))
            self.workshop_table.item(row, 0).setData(Qt.UserRole, id)

    def update_equipment_table(self):
        equipment = self.db.get_equipment()
        self.equipment_table.setRowCount(len(equipment))
        for row, (id, name, article, note) in enumerate(equipment):
            self.equipment_table.setItem(row, 0, QTableWidgetItem(name))
            self.equipment_table.setItem(row, 1, QTableWidgetItem(article or ""))
            self.equipment_table.setItem(row, 2, QTableWidgetItem(note or ""))
            self.equipment_table.item(row, 0).setData(Qt.UserRole, id)

    def add_operation(self):
        dialog = OperationDialog(self, is_edit_db=True, db=self.db)
        if dialog.exec_() == QDialog.Accepted:
            code, name = dialog.get_values()
            if name:
                if self.db.insert_operation(None, "", code, name, ""):
                    self.update_operations_table()
                    QMessageBox.information(self, "Успех", f"'{name}' добавлено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def edit_operation(self):
        row = self.operations_table.currentRow()
        if row >= 0:
            old_code = self.operations_table.item(row, 0).text()
            old_name = self.operations_table.item(row, 1).text()
            dialog = OperationDialog(self, is_edit_db=True, db=self.db)
            dialog.code_combo.setCurrentText(old_code)
            dialog.name_combo.setCurrentText(old_name)
            if dialog.exec_() == QDialog.Accepted:
                code, name = dialog.get_values()
                if name:
                    operation_id = self.operations_table.item(row, 0).data(Qt.UserRole)
                    if self.db.update_operation(operation_id, "", code, name, ""):
                        self.update_operations_table()
                        QMessageBox.information(self, "Успех", "Обновлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def delete_operation(self):
        row = self.operations_table.currentRow()
        if row >= 0:
            name = self.operations_table.item(row, 1).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                operation_id = self.operations_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_operation(operation_id):
                    self.update_operations_table()
                    QMessageBox.information(self, "Успех", f"'{name}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

    def add_workshop(self):
        dialog = WorkshopDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            workshop_name, section, rm = dialog.get_values()
            if workshop_name:
                if self.db.insert_workshop(workshop_name, section, rm):
                    self.update_workshop_table()
                    QMessageBox.information(self, "Успех", f"'{workshop_name}' добавлено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Цех'!")

    def edit_workshop(self):
        row = self.workshop_table.currentRow()
        if row >= 0:
            old_workshop = self.workshop_table.item(row, 0).text()
            old_section = self.workshop_table.item(row, 1).text()
            old_rm = self.workshop_table.item(row, 2).text()
            dialog = WorkshopDialog(self)
            dialog.workshop_input.setText(old_workshop)
            dialog.section_input.setText(old_section)
            dialog.rm_input.setText(old_rm)
            if dialog.exec_() == QDialog.Accepted:
                workshop_name, section, rm = dialog.get_values()
                if workshop_name:
                    workshop_id = self.workshop_table.item(row, 0).data(Qt.UserRole)
                    if self.db.update_workshop(workshop_id, workshop_name, section, rm):
                        self.update_workshop_table()
                        QMessageBox.information(self, "Успех", "Обновлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Цех'!")

    def delete_workshop(self):
        row = self.workshop_table.currentRow()
        if row >= 0:
            workshop_name = self.workshop_table.item(row, 0).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{workshop_name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                workshop_id = self.workshop_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_workshop(workshop_id):
                    self.update_workshop_table()
                    QMessageBox.information(self, "Успех", f"'{workshop_name}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

    def add_equipment(self):
        dialog = EquipmentDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            name, article, note = dialog.get_values()
            if name:
                if self.db.insert_equipment(name, article, note):
                    self.update_equipment_table()
                    QMessageBox.information(self, "Успех", f"'{name}' добавлено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def edit_equipment(self):
        row = self.equipment_table.currentRow()
        if row >= 0:
            old_name = self.equipment_table.item(row, 0).text()
            old_article = self.equipment_table.item(row, 1).text()
            old_note = self.equipment_table.item(row, 2).text()
            dialog = EquipmentDialog(self)
            dialog.name_input.setText(old_name)
            dialog.article_input.setText(old_article)
            dialog.note_input.setText(old_note)
            if dialog.exec_() == QDialog.Accepted:
                name, article, note = dialog.get_values()
                if name:
                    equipment_id = self.equipment_table.item(row, 0).data(Qt.UserRole)
                    if self.db.update_equipment(equipment_id, name, article, note):
                        self.update_equipment_table()
                        QMessageBox.information(self, "Успех", "Обновлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def delete_equipment(self):
        row = self.equipment_table.currentRow()
        if row >= 0:
            name = self.equipment_table.item(row, 0).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                equipment_id = self.equipment_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_equipment(equipment_id):
                    self.update_equipment_table()
                    QMessageBox.information(self, "Успех", f"'{name}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

class EditTPDialog(QDialog):
    def __init__(self, db, parent=None):
        super().__init__(parent)
        self.db = db
        self.setWindowTitle("Редактировать ТП")
        self.setGeometry(100, 100, 800, 600)
        layout = QVBoxLayout()

        self.tab_widget = QTabWidget()
        layout.addWidget(self.tab_widget)

        self.models_tab = QWidget()
        self.tab_widget.addTab(self.models_tab, "Модели")
        self.setup_models_tab()

        self.parts_tab = QWidget()
        self.tab_widget.addTab(self.parts_tab, "Спецификация")
        self.setup_parts_tab()

        self.operations_tab = QWidget()
        self.tab_widget.addTab(self.operations_tab, "Операции")
        self.setup_operations_tab()

        self.document_details_tab = QWidget()
        self.tab_widget.addTab(self.document_details_tab, "Реквизиты документа")
        self.setup_document_details_tab()

        self.update_model_combo()

        buttons = QHBoxLayout()
        import_button = QPushButton("Импорт из Excel")
        import_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        import_button.clicked.connect(self.import_from_excel)
        buttons.addWidget(import_button)
        close_button = QPushButton("Закрыть")
        close_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #FF9800; color: white; border-radius: 5px;")
        close_button.clicked.connect(self.accept)
        buttons.addWidget(close_button)
        layout.addLayout(buttons)
        self.setLayout(layout)

    def setup_models_tab(self):
        layout = QVBoxLayout()
        self.models_table = QTableWidget()
        self.models_table.setColumnCount(1)
        self.models_table.setHorizontalHeaderLabels(['Название'])
        self.models_table.horizontalHeader().setStretchLastSection(True)
        self.models_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.models_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить модель")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_model)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать модель")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_model)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить модель")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_model)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.models_tab.setLayout(layout)
        self.update_models_table()

    def setup_parts_tab(self):
        layout = QVBoxLayout()
        self.parts_model_combo = QComboBox()
        self.parts_model_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addWidget(self.parts_model_combo)
        self.parts_model_combo.currentTextChanged.connect(self.update_parts_table)
        self.parts_table = QTableWidget()
        self.parts_table.setColumnCount(3)
        self.parts_table.setHorizontalHeaderLabels(['Номер', 'Номенклатура', 'Количество'])
        self.parts_table.horizontalHeader().setStretchLastSection(True)
        self.parts_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.parts_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_part)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_part)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_part)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.parts_tab.setLayout(layout)
        self.update_parts_table(self.parts_model_combo.currentText())

    def setup_operations_tab(self):
        layout = QVBoxLayout()
        self.operations_model_combo = QComboBox()
        self.operations_model_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addWidget(self.operations_model_combo)
        self.operations_model_combo.currentTextChanged.connect(self.update_operations_table)
        self.operations_table = QTableWidget()
        self.operations_table.setColumnCount(4)
        self.operations_table.setHorizontalHeaderLabels(['Номер', 'Код', 'Наименование', 'Описание'])
        self.operations_table.horizontalHeader().setStretchLastSection(True)
        self.operations_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.operations_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_operation)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_operation)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_operation)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.operations_tab.setLayout(layout)
        self.update_operations_table(self.operations_model_combo.currentText())

    def setup_document_details_tab(self):
        layout = QVBoxLayout()
        self.document_details_model_combo = QComboBox()
        self.document_details_model_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        layout.addWidget(self.document_details_model_combo)
        self.document_details_model_combo.currentTextChanged.connect(self.update_document_details_table)
        self.document_details_table = QTableWidget()
        self.document_details_table.setColumnCount(5)
        self.document_details_table.setHorizontalHeaderLabels(['Организация', 'Обозначение изделия', 'Обозначение документа', 'Разработал', 'Проверил'])
        self.document_details_table.horizontalHeader().setStretchLastSection(True)
        self.document_details_table.setStyleSheet("font-size: 14px; color: #333;")
        layout.addWidget(self.document_details_table)
        buttons = QHBoxLayout()
        add_button = QPushButton("Добавить")
        add_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #4CAF50; color: white; border-radius: 5px;")
        add_button.clicked.connect(self.add_document_details)
        buttons.addWidget(add_button)
        edit_button = QPushButton("Редактировать")
        edit_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #2196F3; color: white; border-radius: 5px;")
        edit_button.clicked.connect(self.edit_document_details)
        buttons.addWidget(edit_button)
        delete_button = QPushButton("Удалить")
        delete_button.setStyleSheet("font-size: 14px; padding: 8px; background-color: #F44336; color: white; border-radius: 5px;")
        delete_button.clicked.connect(self.delete_document_details)
        buttons.addWidget(delete_button)
        layout.addLayout(buttons)
        self.document_details_tab.setLayout(layout)
        self.update_document_details_table(self.document_details_model_combo.currentText())

    def update_model_combo(self):
        models = [name for _, name in self.db.get_models()]
        for combo in [self.parts_model_combo, self.operations_model_combo, self.document_details_model_combo]:
            combo.clear()
            combo.addItems(models)

    def import_from_excel(self):
        current_model = self.parts_model_combo.currentText()
        if not current_model or not self.db.get_model_id(current_model):
            QMessageBox.warning(self, "Предупреждение", "Сначала добавьте модель в разделе 'Модели'!")
            return
        file_path = QFileDialog.getOpenFileName(self, "Открыть Excel файл", "", "Excel Files (*.xlsx *.xls)")[0]
        if file_path:
            if self.db.import_from_excel(file_path, current_model):
                self.update_parts_table(current_model)
                QMessageBox.information(self, "Успех", "Данные импортированы в Спецификацию!")
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось импортировать данные из Excel. Проверьте консоль для деталей.")

    def update_models_table(self):
        models = self.db.get_models()
        self.models_table.setRowCount(len(models))
        for row, (id, name) in enumerate(models):
            item = QTableWidgetItem(name)
            item.setData(Qt.UserRole, id)
            self.models_table.setItem(row, 0, item)

    def update_parts_table(self, model_name):
        model_id = self.db.get_model_id(model_name)
        parts = self.db.get_parts(model_id) if model_id else []
        self.parts_table.setRowCount(len(parts))
        for row, (id, name, code, quantity) in enumerate(parts):
            self.parts_table.setItem(row, 0, QTableWidgetItem(code or ""))
            self.parts_table.setItem(row, 1, QTableWidgetItem(name))
            self.parts_table.setItem(row, 2, QTableWidgetItem(str(quantity)))
            self.parts_table.item(row, 0).setData(Qt.UserRole, id)

    def update_operations_table(self, model_name):
        model_id = self.db.get_model_id(model_name)
        operations = self.db.get_operations(model_id) if model_id else []
        self.operations_table.setRowCount(len(operations))
        for row, (id, number, code, name, description, _, _, _, _) in enumerate(operations):
            self.operations_table.setItem(row, 0, QTableWidgetItem(number or ""))
            self.operations_table.setItem(row, 1, QTableWidgetItem(code or ""))
            self.operations_table.setItem(row, 2, QTableWidgetItem(name))
            self.operations_table.setItem(row, 3, QTableWidgetItem(description or ""))
            self.operations_table.item(row, 0).setData(Qt.UserRole, id)

    def update_document_details_table(self, model_name):
        model_id = self.db.get_model_id(model_name)
        details = self.db.get_document_details(model_id) if model_id else []
        self.document_details_table.setRowCount(len(details))
        for row, (id, organization, product_code, document_code, developed_by, checked_by) in enumerate(details):
            self.document_details_table.setItem(row, 0, QTableWidgetItem(organization))
            self.document_details_table.setItem(row, 1, QTableWidgetItem(product_code or ""))
            self.document_details_table.setItem(row, 2, QTableWidgetItem(document_code or ""))
            self.document_details_table.setItem(row, 3, QTableWidgetItem(developed_by or ""))
            self.document_details_table.setItem(row, 4, QTableWidgetItem(checked_by or ""))
            self.document_details_table.item(row, 0).setData(Qt.UserRole, id)

    def add_model(self):
        name, ok = QInputDialog.getText(self, "Добавить модель", "Название модели:")
        if ok and name:
            if self.db.insert_model(name):
                self.update_models_table()
                self.update_model_combo()
                if self.parent():
                    self.parent().update_model_combo()
                QMessageBox.information(self, "Успех", f"Модель '{name}' добавлена!")
            else:
                QMessageBox.critical(self, "Ошибка", "Не удалось добавить модель! Возможно, модель с таким именем уже существует.")

    def edit_model(self):
        row = self.models_table.currentRow()
        if row >= 0:
            old_name = self.models_table.item(row, 0).text()
            name, ok = QInputDialog.getText(self, "Редактировать модель", "Название модели:", text=old_name)
            if ok and name:
                model_id = self.models_table.item(row, 0).data(Qt.UserRole)
                if self.db.update_model(model_id, name):
                    self.update_models_table()
                    self.update_model_combo()
                    if self.parent():
                        self.parent().update_model_combo()
                    QMessageBox.information(self, "Успех", f"Модель '{name}' обновлена!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось обновить модель!")

    def delete_model(self):
        row = self.models_table.currentRow()
        if row >= 0:
            name = self.models_table.item(row, 0).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить модель '{name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                model_id = self.models_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_model(model_id):
                    self.update_models_table()
                    self.update_model_combo()
                    if self.parent():
                        self.parent().update_model_combo()
                    QMessageBox.information(self, "Успех", f"Модель '{name}' удалена!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить модель!")

    def add_part(self):
        model_name = self.parts_model_combo.currentText()
        model_id = self.db.get_model_id(model_name)
        if not model_id:
            QMessageBox.critical(self, "Ошибка", "Выберите модель!")
            return
        name, ok = QInputDialog.getText(self, "Добавить деталь", "Название:")
        if ok and name:
            code, ok = QInputDialog.getText(self, "Добавить деталь", "Код:")
            if ok:
                quantity, ok = QInputDialog.getInt(self, "Добавить деталь", "Количество:", min=1)
                if ok:
                    if self.db.insert_part(model_id, name, code, quantity):
                        self.update_parts_table(model_name)
                        QMessageBox.information(self, "Успех", f"'{name}' добавлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")

    def edit_part(self):
        row = self.parts_table.currentRow()
        if row >= 0:
            old_name = self.parts_table.item(row, 1).text()
            name, ok = QInputDialog.getText(self, "Редактировать деталь", "Название:", text=old_name)
            if ok:
                code, ok = QInputDialog.getText(self, "Редактировать деталь", "Код:", text=self.parts_table.item(row, 0).text())
                if ok:
                    quantity, ok = QInputDialog.getInt(self, "Редактировать деталь", "Количество:", value=int(self.parts_table.item(row, 2).text()), min=1)
                    if ok:
                        part_id = self.parts_table.item(row, 0).data(Qt.UserRole)
                        if self.db.update_part(part_id, name, code, quantity):
                            self.update_parts_table(self.parts_model_combo.currentText())
                            QMessageBox.information(self, "Успех", "Обновлено!")
                        else:
                            QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")

    def delete_part(self):
        row = self.parts_table.currentRow()
        if row >= 0:
            name = self.parts_table.item(row, 1).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                part_id = self.parts_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_part(part_id):
                    self.update_parts_table(self.parts_model_combo.currentText())
                    QMessageBox.information(self, "Успех", f"'{name}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

    def add_operation(self):
        model_name = self.operations_model_combo.currentText()
        model_id = self.db.get_model_id(model_name)
        if not model_id:
            QMessageBox.critical(self, "Ошибка", "Выберите модель!")
            return
        dialog = OperationDialog(self, is_edit_db=False, db=self.db)
        if dialog.exec_() == QDialog.Accepted:
            number, code, name, description = dialog.get_values()
            if name:
                if self.db.insert_operation(model_id, number, code, name, description):
                    self.update_operations_table(model_name)
                    QMessageBox.information(self, "Успех", f"'{name}' добавлено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def edit_operation(self):
        row = self.operations_table.currentRow()
        if row >= 0:
            old_number = self.operations_table.item(row, 0).text()
            old_code = self.operations_table.item(row, 1).text()
            old_name = self.operations_table.item(row, 2).text()
            old_description = self.operations_table.item(row, 3).text()
            dialog = OperationDialog(self, is_edit_db=False, db=self.db)
            dialog.number_input.setText(old_number)
            dialog.code_combo.setCurrentText(old_code)
            dialog.name_combo.setCurrentText(old_name)
            dialog.description_input.setText(old_description)
            if dialog.exec_() == QDialog.Accepted:
                number, code, name, description = dialog.get_values()
                if name:
                    operation_id = self.operations_table.item(row, 0).data(Qt.UserRole)
                    if self.db.update_operation(operation_id, number, code, name, description):
                        self.update_operations_table(self.operations_model_combo.currentText())
                        QMessageBox.information(self, "Успех", "Обновлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Наименование'!")

    def delete_operation(self):
        row = self.operations_table.currentRow()
        if row >= 0:
            name = self.operations_table.item(row, 2).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{name}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                operation_id = self.operations_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_operation(operation_id):
                    self.update_operations_table(self.operations_model_combo.currentText())
                    QMessageBox.information(self, "Успех", f"'{name}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

    def add_document_details(self):
        model_name = self.document_details_model_combo.currentText()
        model_id = self.db.get_model_id(model_name)
        if not model_id:
            QMessageBox.critical(self, "Ошибка", "Выберите модель!")
            return
        dialog = DocumentDetailsDialog(self)
        if dialog.exec_() == QDialog.Accepted:
            organization, product_code, document_code, developed_by, checked_by = dialog.get_values()
            if organization:
                if self.db.insert_document_details(model_id, organization, product_code, document_code, developed_by, checked_by):
                    self.update_document_details_table(model_name)
                    QMessageBox.information(self, "Успех", f"'{organization}' добавлено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось добавить!")
            else:
                QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Организация'!")

    def edit_document_details(self):
        row = self.document_details_table.currentRow()
        if row >= 0:
            old_organization = self.document_details_table.item(row, 0).text()
            old_product_code = self.document_details_table.item(row, 1).text()
            old_document_code = self.document_details_table.item(row, 2).text()
            old_developed_by = self.document_details_table.item(row, 3).text()
            old_checked_by = self.document_details_table.item(row, 4).text()
            dialog = DocumentDetailsDialog(self)
            dialog.organization_input.setText(old_organization)
            dialog.product_code_input.setText(old_product_code)
            dialog.document_code_input.setText(old_document_code)
            dialog.developed_by_input.setText(old_developed_by)
            dialog.checked_by_input.setText(old_checked_by)
            if dialog.exec_() == QDialog.Accepted:
                organization, product_code, document_code, developed_by, checked_by = dialog.get_values()
                if organization:
                    details_id = self.document_details_table.item(row, 0).data(Qt.UserRole)
                    if self.db.update_document_details(details_id, organization, product_code, document_code, developed_by, checked_by):
                        self.update_document_details_table(self.document_details_model_combo.currentText())
                        QMessageBox.information(self, "Успех", "Обновлено!")
                    else:
                        QMessageBox.critical(self, "Ошибка", "Не удалось обновить!")
                else:
                    QMessageBox.warning(self, "Предупреждение", "Заполните поле 'Организация'!")

    def delete_document_details(self):
        row = self.document_details_table.currentRow()
        if row >= 0:
            organization = self.document_details_table.item(row, 0).text()
            reply = QMessageBox.question(self, "Подтверждение", f"Удалить '{organization}'?", QMessageBox.Yes | QMessageBox.No)
            if reply == QMessageBox.Yes:
                details_id = self.document_details_table.item(row, 0).data(Qt.UserRole)
                if self.db.delete_document_details(details_id):
                    self.update_document_details_table(self.document_details_model_combo.currentText())
                    QMessageBox.information(self, "Успех", f"'{organization}' удалено!")
                else:
                    QMessageBox.critical(self, "Ошибка", "Не удалось удалить!")

class CAPPWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.db = CAPPDatabase()
        self.setWindowTitle("CAPP Prototype")
        self.setGeometry(100, 100, 800, 600)
        try:
            self.init_ui()
            self.init_data()
        except Exception as e:
            print(f"Ошибка инициализации приложения: {e}")
            traceback.print_exc()
            self.close()

    def init_data(self):
        pass

    def init_ui(self):
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        layout = QVBoxLayout(central_widget)
        layout.setContentsMargins(10, 10, 10, 10)
        layout.setSpacing(10)

        # Кнопки в верхней части
        button_layout = QHBoxLayout()
        generate_btn = QPushButton("Сгенерировать техпроцесс", central_widget)
        generate_btn.setStyleSheet("font-size: 14px; background-color: #808080; color: white; border-radius: 5px;")
        generate_btn.setFixedSize(200, 40)
        generate_btn.clicked.connect(self.generate_process)
        button_layout.addWidget(generate_btn)

        edit_db_btn = QPushButton("Редактировать БД", central_widget)
        edit_db_btn.setStyleSheet("font-size: 14px; background-color: #808080; color: white; border-radius: 5px;")
        edit_db_btn.setFixedSize(200, 40)
        edit_db_btn.clicked.connect(self.edit_db)
        button_layout.addWidget(edit_db_btn)

        edit_tp_btn = QPushButton("Ред. ТП", central_widget)
        edit_tp_btn.setStyleSheet("font-size: 14px; background-color: #808080; color: white; border-radius: 5px;")
        edit_tp_btn.setFixedSize(200, 40)
        edit_tp_btn.clicked.connect(self.edit_tp)
        button_layout.addWidget(edit_tp_btn)

        export_btn = QPushButton("Экспортировать в PDF", central_widget)
        export_btn.setStyleSheet("font-size: 14px; background-color: #808080; color: white; border-radius: 5px;")
        export_btn.setFixedSize(200, 40)
        export_btn.clicked.connect(self.export_to_pdf)
        button_layout.addWidget(export_btn)

        button_layout.addSpacerItem(QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        layout.addLayout(button_layout)

        # Выбор модели
        input_layout = QHBoxLayout()
        model_label = QLabel("Модель:", central_widget)
        model_label.setStyleSheet("font-size: 14px; font-weight: bold; color: #333;")
        model_label.setFixedWidth(100)
        input_layout.addWidget(model_label)
        self.model_combo = QComboBox(central_widget)
        self.model_combo.setStyleSheet("font-size: 14px; padding: 5px;")
        input_layout.addWidget(self.model_combo)
        input_layout.addSpacerItem(QSpacerItem(20, 0, QSizePolicy.Expanding, QSizePolicy.Minimum))
        self.update_model_combo()
        layout.addLayout(input_layout)

        # Разделитель
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # Группа для модели
        model_group = QGroupBox("Модель")
        model_group.setStyleSheet("QGroupBox { font-size: 16px; font-weight: bold; color: #2E7D32; }")
        model_group.setFlat(False)
        model_layout = QVBoxLayout()
        self.model_label = QLabel("")
        self.model_label.setStyleSheet("font-size: 14px; padding: 5px; color: #333;")
        model_layout.addWidget(self.model_label)
        model_group.setLayout(model_layout)
        layout.addWidget(model_group)

        # Группа для спецификации
        parts_group = QGroupBox("Спецификация")
        parts_group.setStyleSheet("QGroupBox { font-size: 16px; font-weight: bold; color: #2E7D32; }")
        parts_group.setFlat(False)
        parts_group.setCheckable(True)
        parts_group.setChecked(True)
        parts_layout = QVBoxLayout()
        self.parts_table = QTableWidget()
        self.parts_table.setColumnCount(3)
        self.parts_table.setHorizontalHeaderLabels(['Номер', 'Номенклатура', 'Количество'])
        self.parts_table.horizontalHeader().setStretchLastSection(True)
        self.parts_table.setStyleSheet("font-size: 14px; color: #333;")
        parts_scroll = QScrollArea()
        parts_scroll.setWidgetResizable(True)
        parts_scroll.setWidget(self.parts_table)
        parts_layout.addWidget(parts_scroll)
        parts_group.setLayout(parts_layout)
        layout.addWidget(parts_group)

        # Разделитель
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # Группа для операций
        operations_group = QGroupBox("Операции")
        operations_group.setStyleSheet("QGroupBox { font-size: 16px; font-weight: bold; color: #2E7D32; }")
        operations_group.setFlat(False)
        operations_group.setCheckable(True)
        operations_group.setChecked(True)
        operations_layout = QVBoxLayout()
        self.operations_table = QTableWidget()
        self.operations_table.setColumnCount(4)
        self.operations_table.setHorizontalHeaderLabels(['Номер', 'Код', 'Наименование', 'Оборудование'])
        self.operations_table.horizontalHeader().setStretchLastSection(True)
        self.operations_table.setStyleSheet("font-size: 14px; color: #333;")
        operations_scroll = QScrollArea()
        operations_scroll.setWidgetResizable(True)
        operations_scroll.setWidget(self.operations_table)
        operations_layout.addWidget(operations_scroll)
        operations_group.setLayout(operations_layout)
        layout.addWidget(operations_group)

        # Разделитель
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # Группа для расцеховки
        workshop_group = QGroupBox("Расцеховка")
        workshop_group.setStyleSheet("QGroupBox { font-size: 16px; font-weight: bold; color: #2E7D32; }")
        workshop_group.setFlat(False)
        workshop_group.setCheckable(True)
        workshop_group.setChecked(True)
        workshop_layout = QVBoxLayout()
        self.workshop_table = QTableWidget()
        self.workshop_table.setColumnCount(3)
        self.workshop_table.setHorizontalHeaderLabels(['Цех', 'Участок', 'РМ'])
        self.workshop_table.horizontalHeader().setStretchLastSection(True)
        self.workshop_table.setStyleSheet("font-size: 14px; color: #333;")
        workshop_scroll = QScrollArea()
        workshop_scroll.setWidgetResizable(True)
        workshop_scroll.setWidget(self.workshop_table)
        workshop_layout.addWidget(workshop_scroll)
        workshop_group.setLayout(workshop_layout)
        layout.addWidget(workshop_group)

        # Разделитель
        separator = QFrame()
        separator.setFrameShape(QFrame.HLine)
        separator.setFrameShadow(QFrame.Sunken)
        layout.addWidget(separator)

        # Группа для оборудования
        equipment_group = QGroupBox("Оборудование")
        equipment_group.setStyleSheet("QGroupBox { font-size: 16px; font-weight: bold; color: #2E7D32; }")
        equipment_group.setFlat(False)
        equipment_group.setCheckable(True)
        equipment_group.setChecked(True)
        equipment_layout = QVBoxLayout()
        self.equipment_table = QTableWidget()
        self.equipment_table.setColumnCount(3)
        self.equipment_table.setHorizontalHeaderLabels(['Наименование', 'Артикул', 'Примечание'])
        self.equipment_table.horizontalHeader().setStretchLastSection(True)
        self.equipment_table.setStyleSheet("font-size: 14px; color: #333;")
        equipment_scroll = QScrollArea()
        equipment_scroll.setWidgetResizable(True)
        equipment_scroll.setWidget(self.equipment_table)
        equipment_layout.addWidget(equipment_scroll)
        equipment_group.setLayout(equipment_layout)
        layout.addWidget(equipment_group)

        # Добавляем растягивающийся разделитель
        layout.addSpacerItem(QSpacerItem(20, 40, QSizePolicy.Minimum, QSizePolicy.Expanding))

    def update_model_combo(self):
        self.model_combo.clear()
        self.model_combo.addItems([name for _, name in self.db.get_models()])

    def edit_db(self):
        dialog = EditDBDialog(self.db, self)
        dialog.exec_()

    def edit_tp(self):
        dialog = EditTPDialog(self.db, self)
        dialog.exec_()
        self.update_model_combo()

    def generate_process(self):
        model = self.model_combo.currentText()
        if not model:
            QMessageBox.warning(self, "Ошибка", "Выберите модель!")
            return

        model_id = self.db.get_model_id(model)
        if not model_id:
            QMessageBox.critical(self, "Ошибка", f"Модель {model} не найдена в базе данных!")
            return

        parts = self.db.get_parts(model_id)
        operations = self.db.get_operations(model_id)
        workshops = self.db.get_workshop()
        equipment = self.db.get_equipment()
        document_details = self.db.get_document_details(model_id)

        self.model_label.setText(f"Модель: {model}")

        self.parts_table.setRowCount(len(parts))
        for row, (id, name, code, quantity) in enumerate(parts):
            self.parts_table.setItem(row, 0, QTableWidgetItem(code or ""))
            self.parts_table.setItem(row, 1, QTableWidgetItem(name))
            self.parts_table.setItem(row, 2, QTableWidgetItem(str(quantity)))

        self.operations_table.setRowCount(len(operations))
        for row, (id, number, code, name, _, equipment_name, _, _, _) in enumerate(operations):
            self.operations_table.setItem(row, 0, QTableWidgetItem(number or ""))
            self.operations_table.setItem(row, 1, QTableWidgetItem(code or ""))
            self.operations_table.setItem(row, 2, QTableWidgetItem(name))
            self.operations_table.setItem(row, 3, QTableWidgetItem(equipment_name or ""))

        self.workshop_table.setRowCount(len(workshops))
        for row, (id, workshop_name, section, rm) in enumerate(workshops):
            self.workshop_table.setItem(row, 0, QTableWidgetItem(workshop_name))
            self.workshop_table.setItem(row, 1, QTableWidgetItem(section or ""))
            self.workshop_table.setItem(row, 2, QTableWidgetItem(rm or ""))

        self.equipment_table.setRowCount(len(equipment))
        for row, (id, name, article, note) in enumerate(equipment):
            self.equipment_table.setItem(row, 0, QTableWidgetItem(name))
            self.equipment_table.setItem(row, 1, QTableWidgetItem(article or ""))
            self.equipment_table.setItem(row, 2, QTableWidgetItem(note or ""))

        self.process_data = {
            'model': model,
            'parts': parts,
            'operations': operations,
            'workshops': workshops,
            'equipment': equipment,
            'document_details': document_details,
            'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }

    def export_to_pdf(self):
        if not hasattr(self, 'process_data'):
            QMessageBox.warning(self, "Ошибка", "Сначала сгенерируйте техпроцесс!")
            return

        file_path = QFileDialog.getSaveFileName(self, "Сохранить PDF",
                                                f"TechProcess_{self.process_data['model']}_{self.process_data['timestamp'].replace(':', '-')}.pdf",
                                                "PDF Files (*.pdf)")[0]
        if not file_path:
            return

        try:
            if hasattr(sys, '_MEIPASS'):
                font_path = os.path.join(sys._MEIPASS, 'DejaVuSans.ttf')
            else:
                font_path = os.path.join(os.getcwd(), 'DejaVuSans.ttf')
            print(f"Попытка загрузки шрифта: {font_path}")
            if not os.path.exists(font_path):
                print(f"Ошибка: Файл шрифта {font_path} не найден, используется Helvetica")
                font = "Helvetica"
            else:
                pdfmetrics.registerFont(TTFont('DejaVuSans', font_path))
                font = 'DejaVuSans'
                print("Шрифт DejaVuSans успешно зарегистрирован")

            c = canvas.Canvas(file_path, pagesize=A4)
            c.setFont(font, 12)
            y = 280 * mm

            c.drawString(20 * mm, y, f"Модель: {self.process_data['model']}")
            y -= 15 * mm

            c.drawString(20 * mm, y, "Реквизиты документа:")
            y -= 10 * mm
            for _, organization, product_code, document_code, developed_by, checked_by in self.process_data['document_details']:
                c.drawString(20 * mm, y, f"Организация: {organization}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Обозначение изделия: {product_code or 'Не указано'}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Обозначение документа: {document_code or 'Не указано'}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Разработал: {developed_by or 'Не указано'}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Проверил: {checked_by or 'Не указано'}")
                y -= 10 * mm

            c.drawString(20 * mm, y, "Расцеховка:")
            y -= 10 * mm
            for _, workshop_name, section, rm in self.process_data['workshops']:
                c.drawString(20 * mm, y, f"Цех: {workshop_name}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Участок/РМ: {section or ''}/{rm or ''}")
                y -= 10 * mm

            c.drawString(20 * mm, y, "Операции:")
            y -= 10 * mm
            for _, number, _, name, _, equipment, _, _, _ in self.process_data['operations']:
                c.drawString(20 * mm, y, f"{number} - {name}")
                y -= 10 * mm
                c.drawString(25 * mm, y, f"Оборудование: {equipment or 'Не указано'}")
                y -= 10 * mm

            y -= 10 * mm
            c.drawString(20 * mm, y, "Спецификация:")
            y -= 10 * mm
            for _, name, code, quantity in self.process_data['parts']:
                c.drawString(20 * mm, y, f"- {name} ({code or 'Не указан'}): {quantity}")
                y -= 10 * mm

            y -= 10 * mm
            c.drawString(20 * mm, y, "Оборудование:")
            y -= 10 * mm
            for _, name, article, note in self.process_data['equipment']:
                c.drawString(20 * mm, y, f"- {name} (Артикул: {article or 'Не указан'}, Примечание: {note or 'Не указано'})")
                y -= 10 * mm

            if y < 20 * mm:
                c.showPage()
                c.setFont(font, 12)
                y = 280 * mm

            c.drawString(20 * mm, y, f"Дата: {self.process_data['timestamp']}")
            c.save()
            print(f"PDF успешно сохранен: {file_path}")
            QMessageBox.information(self, "Успех", f"PDF успешно сохранен: {file_path}")
        except Exception as e:
            print(f"Ошибка экспорта в PDF: {e}")
            QMessageBox.critical(self, "Ошибка", f"Не удалось экспортировать в PDF: {e}")

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = CAPPWindow()
    window.show()
    sys.exit(app.exec_())