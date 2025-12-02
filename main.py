# main.py
import sys
import os
import tempfile
import traceback
import pandas as pd
import xlwings as xw
from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QMessageBox


class ExcelMerger(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()

        self.setWindowTitle("Excel Merger PyQt5 (.xls + .xlsx)")
        self.setGeometry(200, 200, 720, 380)

        self.files = []
        self.output_file = ""
        self.temp_files = []

        self._build_ui()

    def _build_ui(self):
        layout = QtWidgets.QVBoxLayout()

        # Folder / files selection
        row1 = QtWidgets.QHBoxLayout()
        self.folder_input = QtWidgets.QLineEdit()
        btn_folder = QtWidgets.QPushButton("Выбрать папку")
        btn_folder.clicked.connect(self.select_folder)
        row1.addWidget(QtWidgets.QLabel("Папка (необязательно):"))
        row1.addWidget(self.folder_input)
        row1.addWidget(btn_folder)
        layout.addLayout(row1)

        row2 = QtWidgets.QHBoxLayout()
        self.files_input = QtWidgets.QLineEdit()
        btn_files = QtWidgets.QPushButton("Выбрать файлы")
        btn_files.clicked.connect(self.select_files)
        row2.addWidget(QtWidgets.QLabel("Или файлы:"))
        row2.addWidget(self.files_input)
        row2.addWidget(btn_files)
        layout.addLayout(row2)

        # Output
        row3 = QtWidgets.QHBoxLayout()
        self.output_input = QtWidgets.QLineEdit("merged.xlsx")
        btn_output = QtWidgets.QPushButton("Выбрать итоговый файл")
        btn_output.clicked.connect(self.select_output)
        row3.addWidget(QtWidgets.QLabel("Имя итогового файла:"))
        row3.addWidget(self.output_input)
        row3.addWidget(btn_output)
        layout.addLayout(row3)

        # Mode
        layout.addWidget(QtWidgets.QLabel("Режим объединения:"))
        self.mode_sep = QtWidgets.QRadioButton("Каждый лист → отдельный лист")
        self.mode_sep.setChecked(True)
        self.mode_one = QtWidgets.QRadioButton("Все данные → одна таблица (один под один)")
        layout.addWidget(self.mode_sep)
        layout.addWidget(self.mode_one)

        # Progress
        self.progress = QtWidgets.QProgressBar()
        layout.addWidget(self.progress)

        # Buttons
        btn_row = QtWidgets.QHBoxLayout()
        btn_merge = QtWidgets.QPushButton("Объединить")
        btn_merge.clicked.connect(self.on_merge)
        btn_exit = QtWidgets.QPushButton("Выход")
        btn_exit.clicked.connect(self.close)
        btn_row.addWidget(btn_merge)
        btn_row.addWidget(btn_exit)
        layout.addLayout(btn_row)

        self.setLayout(layout)

    # ---------- selection handlers ----------
    def select_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if folder:
            self.folder_input.setText(folder)

    def select_files(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Выбрать Excel файлы", "", "Excel Files (*.xls *.xlsx)"
        )
        if files:
            self.files = files
            self.files_input.setText(";".join([os.path.basename(f) for f in files]))

    def select_output(self):
        file, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Сохранить как", "", "Excel Files (*.xlsx)"
        )
        if file:
            if not file.lower().endswith(".xlsx"):
                file += ".xlsx"
            self.output_input.setText(file)
            self.output_file = file

    # ---------- main flow ----------
    def on_merge(self):
        try:
            folder = self.folder_input.text().strip()
            selected_files_field = self.files_input.text().strip()
            files = []

            if selected_files_field:
                # files_input may contain basenames only if chosen earlier; prefer self.files if populated
                if self.files:
                    files.extend(self.files)
                else:
                    # split by ; (from text field)
                    files += [p.strip() for p in selected_files_field.split(";") if p.strip()]

            if folder:
                # add all xls/xlsx from folder
                for name in os.listdir(folder):
                    if name.lower().endswith((".xls", ".xlsx")):
                        files.append(os.path.join(folder, name))

            # remove duplicates, preserve order
            seen = set()
            files_unique = []
            for f in files:
                if f not in seen:
                    seen.add(f)
                    files_unique.append(f)
            files = files_unique

            if not files:
                QMessageBox.warning(self, "Ошибка", "Не выбраны файлы и не указана папка.")
                return

            out_name = self.output_input.text().strip()
            if not out_name:
                QMessageBox.warning(self, "Ошибка", "Укажите имя итогового файла.")
                return
            if not out_name.lower().endswith(".xlsx"):
                out_name += ".xlsx"
            self.output_file = out_name

            # reset progress and temp list
            self.progress.setValue(0)
            QApplication = QtWidgets.QApplication.instance()
            QApplication.processEvents()
            self.temp_files = []

            # Step 1: convert .xls -> .xlsx using single hidden xlwings App
            self._convert_xls_to_xlsx(files)

            # Step 2: merge according to mode
            if self.mode_sep.isChecked():
                self._merge_each_sheet_separate(self.temp_files)
            else:
                self._merge_into_one_table(self.temp_files)

            # Step 3: cleanup temp files
            self._cleanup_temp_files()

            QMessageBox.information(self, "Готово", f"Файл сохранён:\n{self.output_file}")
            self.progress.setValue(100)

        except Exception as e:
            tb = traceback.format_exc()
            QMessageBox.critical(self, "Ошибка", f"Произошла ошибка:\n{e}\n\n{tb}")

    # ---------- helpers ----------
    def _convert_xls_to_xlsx(self, files):
        total = len(files)
        # create a single hidden Excel app
        app = None
        try:
            app = xw.App(visible=False, add_book=False)
        except Exception as e:
            # xlwings may fail when Excel not installed; if so, try skipping conversion and rely on pandas/xlrd if available
            app = None

        for idx, f in enumerate(files, 1):
            lower = f.lower()
            if lower.endswith(".xls") and app is not None:
                try:
                    # create unique temp file path
                    tf = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    tf.close()
                    temp_path = tf.name
                    wb = app.books.open(f)
                    wb.api.DisplayAlerts = False
                    wb.save(temp_path)
                    wb.close()
                    self.temp_files.append(temp_path)
                except Exception as e:
                    # fallback: try using pandas read + to_excel (less reliable for complex xls)
                    try:
                        df_all = pd.read_excel(f, sheet_name=None)  # dict of dfs
                        # write to a temp xlsx with multiple sheets
                        temp_dir = tempfile.mkdtemp()
                        temp_path = os.path.join(temp_dir, os.path.basename(f) + ".xlsx")
                        with pd.ExcelWriter(temp_path, engine="openpyxl") as w:
                            for sheetn, dff in df_all.items():
                                dff.to_excel(w, sheet_name=str(sheetn)[:31], index=False)
                        self.temp_files.append(temp_path)
                    except Exception as e2:
                        QMessageBox.warning(self, "Ошибка", f"Не удалось сконвертировать {f}:\n{e}\n{e2}")
                finally:
                    # update progress
                    self.progress.setValue(int(idx / total * 40))
                    QtWidgets.QApplication.processEvents()
            else:
                # already xlsx (or xlwings not available) — just append original
                self.temp_files.append(f)
                self.progress.setValue(int(idx / total * 40))
                QtWidgets.QApplication.processEvents()

        # close app if created
        try:
            if app is not None:
                app.quit()
        except:
            pass

    def _merge_each_sheet_separate(self, files):
        # write each sheet to its own unique sheet with header line (full file name) above table
        writer = pd.ExcelWriter(self.output_file, engine="openpyxl")
        sheet_index = 1
        total = len(files)
        for i, file in enumerate(files, 1):
            try:
                # read all sheets
                excel = pd.ExcelFile(file, engine="openpyxl")
                base_name = os.path.splitext(os.path.basename(file))[0]
                for sheet in excel.sheet_names:
                    df = excel.parse(sheet)

                    # create header rows: first row contains full filename
                    header_df = pd.DataFrame([[f"Полное имя файла: {os.path.basename(file)}"]])
                    empty_row = pd.DataFrame([[]])
                    final_df = pd.concat([header_df, empty_row, df], ignore_index=True)

                    # decide sheet name: try to keep readable but unique and <=31 chars
                    safe_file = base_name[:20]
                    safe_sheet = sheet[:20]
                    candidate = f"{sheet_index}_{safe_file}_{safe_sheet}"
                    candidate = candidate[:31]
                    # ensure uniqueness by appending counter until unique
                    while candidate in writer.book.sheetnames:
                        sheet_index += 1
                        candidate = f"{sheet_index}_{safe_file}_{safe_sheet}"
                        candidate = candidate[:31]

                    # write without headers because we already included them in final_df
                    final_df.to_excel(writer, sheet_name=candidate, index=False, header=False)
                    sheet_index += 1

            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось обработать {file}:\n{e}")
            # progress: 40..90
            self.progress.setValue(int(i / total * 50 + 40))
            QtWidgets.QApplication.processEvents()
        writer.close()

    def _merge_into_one_table(self, files):
        total = len(files)
        all_blocks = []
        counter = 1
        for i, file in enumerate(files, 1):
            try:
                excel = pd.ExcelFile(file, engine="openpyxl")
                for sheet in excel.sheet_names:
                    df = excel.parse(sheet)

                    # columns with source
                    df.insert(0, "Source_Sheet", sheet)
                    df.insert(0, "Source_File", os.path.basename(file))

                    # If sheet name is long (>20), convert to Sheet_N and also insert original name
                    if len(sheet) > 20:
                        # insert original name column as first column to be visible
                        df.insert(0, "Sheet_Name_Original", sheet)
                        df["Source_Sheet"] = f"Sheet_{counter}"
                        counter += 1

                    # build block with header rows: full filename line + empty line + df
                    header_df = pd.DataFrame([[f"Полное имя файла: {os.path.basename(file)}"]])
                    empty_row = pd.DataFrame([[]])
                    block = pd.concat([header_df, empty_row, df], ignore_index=True)
                    all_blocks.append(block)
            except Exception as e:
                QMessageBox.warning(self, "Ошибка", f"Не удалось обработать {file}:\n{e}")

            # progress: 40..90
            self.progress.setValue(int(i / total * 50 + 40))
            QtWidgets.QApplication.processEvents()

        # concat all blocks and save without header (because blocks already include header rows)
        if all_blocks:
            final = pd.concat(all_blocks, ignore_index=True)
            # write with header=False so that the inserted header lines remain as-is
            final.to_excel(self.output_file, index=False, header=False)

    def _cleanup_temp_files(self):
        for f in self.temp_files:
            try:
                if os.path.exists(f):
                    os.remove(f)
            except Exception:
                # просто игнорируем ошибку удаления
                pass
        self.temp_files = []


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = ExcelMerger()
    w.show()
    sys.exit(app.exec_())


if __name__ == "__main__":
    main()