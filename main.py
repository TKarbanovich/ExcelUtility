import sys
import os
import pandas as pd
import xlwings as xw
from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QRadioButton,
    QButtonGroup, QProgressBar, QMessageBox
)


class ExcelMerger(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merger PyQt5 (.xls + .xlsx)")
        self.setGeometry(200, 200, 650, 380)
        self.selected_files = []
        self.init_ui()

    # ==========================================================
    # UI
    # ==========================================================
    def init_ui(self):
        layout = QVBoxLayout()

        # Папка
        folder_layout = QHBoxLayout()
        self.folder_input = QLineEdit()
        folder_btn = QPushButton("Выбрать папку")
        folder_btn.clicked.connect(self.choose_folder)
        folder_layout.addWidget(QLabel("Папка (необязательно):"))
        folder_layout.addWidget(self.folder_input)
        folder_layout.addWidget(folder_btn)
        layout.addLayout(folder_layout)

        # Файлы
        files_layout = QHBoxLayout()
        self.files_input = QLineEdit()
        files_btn = QPushButton("Выбрать файлы")
        files_btn.clicked.connect(self.choose_files)
        files_layout.addWidget(QLabel("Выбор файлов:"))
        files_layout.addWidget(self.files_input)
        files_layout.addWidget(files_btn)
        layout.addLayout(files_layout)

        # Имя итогового файла
        out_layout = QHBoxLayout()
        self.output_input = QLineEdit("merged.xlsx")
        out_layout.addWidget(QLabel("Имя итогового файла:"))
        out_layout.addWidget(self.output_input)
        layout.addLayout(out_layout)

        # Режим
        self.mode1_radio = QRadioButton("Каждый лист → отдельный лист")
        self.mode1_radio.setChecked(True)
        self.mode2_radio = QRadioButton("Все таблицы → одна таблица")

        group = QButtonGroup()
        group.addButton(self.mode1_radio)
        group.addButton(self.mode2_radio)

        layout.addWidget(self.mode1_radio)
        layout.addWidget(self.mode2_radio)

        # Прогресс
        self.progress = QProgressBar()
        layout.addWidget(self.progress)

        # Кнопки
        btn_layout = QHBoxLayout()
        merge_btn = QPushButton("Объединить")
        merge_btn.clicked.connect(self.start_merge)
        exit_btn = QPushButton("Выход")
        exit_btn.clicked.connect(self.close)
        btn_layout.addWidget(merge_btn)
        btn_layout.addWidget(exit_btn)
        layout.addLayout(btn_layout)

        self.setLayout(layout)

    # ==========================================================
    # Выбор папки / файлов
    # ==========================================================
    def choose_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if folder:
            self.folder_input.setText(folder)

    def choose_files(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Выбрать Excel файлы", "", "Excel Files (*.xls *.xlsx)"
        )
        if files:
            self.selected_files = files
            self.files_input.setText("; ".join(os.path.basename(f) for f in files))

    # ==========================================================
    # Запуск объединения
    # ==========================================================
    def start_merge(self):
        folder = self.folder_input.text().strip()
        files = self.selected_files.copy()

        if folder:
            for f in os.listdir(folder):
                if f.lower().endswith((".xls", ".xlsx")):
                    files.append(os.path.join(folder, f))

        if not files:
            QMessageBox.warning(self, "Ошибка", "Нет выбранных файлов.")
            return

        out_name = self.output_input.text().strip()
        if not out_name.endswith(".xlsx"):
            out_name += ".xlsx"
        output_path = os.path.join(os.getcwd(), out_name)

        self.progress.setValue(0)
        QApplication.processEvents()

        # ------------------------------------------------------
        # Конвертация XLS → XLSX скрытым Excel
        # ------------------------------------------------------
        app = xw.App(visible=False, add_book=False)

        temp_files = []
        for i, f in enumerate(files, 1):
            if f.lower().endswith(".xls"):
                try:
                    wb = app.books.open(f)
                    new_path = f + "x"  # file.xls → file.xlsx
                    wb.save(new_path)
                    wb.close()
                    temp_files.append(new_path)
                except Exception as e:
                    QMessageBox.warning(self, "Ошибка", f"Ошибка при конвертации {f}:\n{e}")
            else:
                temp_files.append(f)

            self.progress.setValue(int(i / len(files) * 40))
            QApplication.processEvents()

        app.quit()  # гарантированно закрываем Excel

        # ------------------------------------------------------
        # Объединение
        # ------------------------------------------------------
        if self.mode1_radio.isChecked():
            self.merge_separate_sheets(temp_files, output_path)
        else:
            self.merge_into_one(temp_files, output_path)

        # Удаляем временные xlsx
        for f in temp_files:
            if f.endswith(".xlsx") and f[:-1].endswith(".xls"):
                try:
                    os.remove(f)
                except:
                    pass

        QMessageBox.information(self, "Готово", f"Файл сохранён:\n{output_path}")
        self.progress.setValue(100)

    # ==========================================================
    # Режим 1: каждый лист → в отдельный лист
    # ==========================================================
    def merge_separate_sheets(self, files, output_path):
        writer = pd.ExcelWriter(output_path, engine="openpyxl")
        sheet_index = 1
        total = len(files)

        for i, file in enumerate(files, 1):
            try:
                excel = pd.ExcelFile(file, engine="openpyxl")
                base = os.path.splitext(os.path.basename(file))[0]

                for sheet in excel.sheet_names:
                    df = excel.parse(sheet)

                    # формируем 100% уникальное имя листа
                    clean_file = base[:20]
                    clean_sheet = sheet[:20]
                    name = f"{sheet_index}_{clean_file}_{clean_sheet}"

                    name = name[:31]  # ограничение Excel
                    sheet_index += 1

                    df.to_excel(writer, sheet_name=name, index=False)

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Ошибка в файле {file}:\n{e}")

            self.progress.setValue(int(i / total * 60 + 40))
            QApplication.processEvents()

        writer.close()

    # ==========================================================
    # Режим 2: все листы → в одну таблицу
    # ==========================================================
    def merge_into_one(self, files, output_path):
        all_dfs = []
        total = len(files)
        counter = 1

        for i, file in enumerate(files, 1):
            try:
                excel = pd.ExcelFile(file, engine="openpyxl")
                for sheet in excel.sheet_names:
                    df = excel.parse(sheet)

                    df["Source_File"] = os.path.basename(file)
                    df["Source_Sheet"] = sheet

                    # длинные названия → приводим
                    if len(sheet) > 20:
                        df.insert(0, "Sheet_Name_Original", sheet)
                        df["Source_Sheet"] = f"Sheet_{counter}"
                        counter += 1

                    all_dfs.append(df)

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Ошибка в файле {file}:\n{e}")

            self.progress.setValue(int(i / total * 60 + 40))
            QApplication.processEvents()

        if all_dfs:
            final_df = pd.concat(all_dfs, ignore_index=True)
            final_df.to_excel(output_path, index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = ExcelMerger()
    window.show()
    sys.exit(app.exec_())