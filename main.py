# main.py
import sys
import os
import tempfile
import traceback
import pandas as pd
import xlwings as xw
from PyQt5 import QtWidgets, QtCore
from PyQt5.QtWidgets import QMessageBox

# ------- MAIN WIDGET -------
class ExcelMerger(QtWidgets.QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Excel Merger (PyQt5)")
        self.setGeometry(300, 200, 760, 380)

        self.files = []             # list of original file paths chosen by user
        self.output_file = ""       # output .xlsx path
        self.mapping = []           # list of tuples (original_path, work_path)
        self.temp_files = []        # list of generated temp xlsx files

        self._build_ui()

    def _build_ui(self):
        v = QtWidgets.QVBoxLayout()

        # folder / files
        h1 = QtWidgets.QHBoxLayout()
        self.folder_edit = QtWidgets.QLineEdit()
        btn_folder = QtWidgets.QPushButton("Выбрать папку")
        btn_folder.clicked.connect(self.choose_folder)
        h1.addWidget(QtWidgets.QLabel("Папка (необязательно):"))
        h1.addWidget(self.folder_edit)
        h1.addWidget(btn_folder)
        v.addLayout(h1)

        h2 = QtWidgets.QHBoxLayout()
        self.files_edit = QtWidgets.QLineEdit()
        btn_files = QtWidgets.QPushButton("Выбрать файлы")
        btn_files.clicked.connect(self.choose_files)
        h2.addWidget(QtWidgets.QLabel("Или файлы:"))
        h2.addWidget(self.files_edit)
        h2.addWidget(btn_files)
        v.addLayout(h2)

        # output
        h3 = QtWidgets.QHBoxLayout()
        self.output_edit = QtWidgets.QLineEdit("merged.xlsx")
        btn_output = QtWidgets.QPushButton("Выбрать итоговый файл")
        btn_output.clicked.connect(self.choose_output)
        h3.addWidget(QtWidgets.QLabel("Имя итогового файла:"))
        h3.addWidget(self.output_edit)
        h3.addWidget(btn_output)
        v.addLayout(h3)

        # mode
        v.addWidget(QtWidgets.QLabel("Режим объединения:"))
        self.mode_sep = QtWidgets.QRadioButton("Каждый лист → отдельный лист")
        self.mode_sep.setChecked(True)
        self.mode_one = QtWidgets.QRadioButton("В одну таблицу (один под один)")
        v.addWidget(self.mode_sep)
        v.addWidget(self.mode_one)

        # progress
        self.progress = QtWidgets.QProgressBar()
        self.progress.setRange(0, 100)
        v.addWidget(self.progress)

        # buttons
        hbtn = QtWidgets.QHBoxLayout()
        btn_run = QtWidgets.QPushButton("Объединить")
        btn_run.clicked.connect(self.on_run)
        btn_cancel = QtWidgets.QPushButton("Выход")
        btn_cancel.clicked.connect(self.close)
        hbtn.addWidget(btn_run)
        hbtn.addWidget(btn_cancel)
        v.addLayout(hbtn)

        self.setLayout(v)

    # ---------- selectors ----------
    def choose_folder(self):
        folder = QtWidgets.QFileDialog.getExistingDirectory(self, "Выбрать папку")
        if folder:
            self.folder_edit.setText(folder)
            # populate files list from folder
            files = [
                os.path.join(folder, f)
                for f in os.listdir(folder)
                if f.lower().endswith((".xls", ".xlsx"))
            ]
            self.files = files
            self.files_edit.setText(";".join([os.path.basename(f) for f in files]))

    def choose_files(self):
        files, _ = QtWidgets.QFileDialog.getOpenFileNames(
            self, "Выбрать Excel файлы", "", "Excel Files (*.xls *.xlsx)"
        )
        if files:
            self.files = files
            self.files_edit.setText(";".join([os.path.basename(f) for f in files]))

    def choose_output(self):
        file, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Сохранить как", "", "Excel Files (*.xlsx)"
        )
        if file:
            if not file.lower().endswith(".xlsx"):
                file += ".xlsx"
            self.output_edit.setText(file)
            self.output_file = file

    # ---------- main ----------
    def on_run(self):
        try:
            # prepare files list: prefer explicit list (self.files), or folder input
            files = []
            if self.files:
                files = list(self.files)
            else:
                folder = self.folder_edit.text().strip()
                if folder:
                    for name in os.listdir(folder):
                        if name.lower().endswith((".xls", ".xlsx")):
                            files.append(os.path.join(folder, name))

            # remove duplicates
            seen = set()
            files_unique = []
            for p in files:
                if p not in seen:
                    seen.add(p)
                    files_unique.append(p)
            files = files_unique

            if not files:
                QMessageBox.warning(self, "Ошибка", "Не выбраны файлы и не указана папка.")
                return

            out_name = self.output_edit.text().strip()
            if not out_name:
                QMessageBox.warning(self, "Ошибка", "Укажите имя итогового файла.")
                return
            if not out_name.lower().endswith(".xlsx"):
                out_name += ".xlsx"
            self.output_file = out_name

            self.progress.setValue(0)
            QtWidgets.QApplication.processEvents()

            # Step A: build mapping original -> work (convert .xls to temp .xlsx)
            self.mapping = []
            self.temp_files = []
            self._convert_xls_to_xlsx(files)

            # Step B: merge according to mode
            if self.mode_sep.isChecked():
                self._merge_each_sheet_separate(self.mapping)
            else:
                self._merge_into_one_table(self.mapping)

            # Step C: cleanup temp files
            self._cleanup_temp_files()

            QMessageBox.information(self, "Готово", f"Файл сохранён:\n{self.output_file}")
            self.progress.setValue(100)

        except Exception as e:
            tb = traceback.format_exc()
            QMessageBox.critical(self, "Ошибка", f"{e}\n\n{tb}")

    # ---------- conversion ----------
    def _convert_xls_to_xlsx(self, files):
        total = len(files)
        app = None
        try:
            # create single hidden Excel app
            app = xw.App(visible=False, add_book=False)
        except Exception:
            app = None  # xlwings not available or Excel not installed -> we'll try pandas fallback

        for idx, orig in enumerate(files, 1):
            lower = orig.lower()
            if lower.endswith(".xls") and app is not None:
                try:
                    # create temp file
                    tf = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                    tf.close()
                    temp_path = tf.name
                    wb = app.books.open(orig)
                    # suppress alerts
                    try:
                        wb.api.DisplayAlerts = False
                    except Exception:
                        pass
                    wb.save(temp_path)
                    wb.close()
                    work = temp_path
                    self.mapping.append((orig, work))
                    self.temp_files.append(temp_path)
                except Exception as e:
                    # fallback: try pandas read/write (less reliable for some .xls)
                    try:
                        all_sheets = pd.read_excel(orig, sheet_name=None)
                        tf = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
                        tf.close()
                        temp_path = tf.name
                        with pd.ExcelWriter(temp_path, engine="openpyxl") as w:
                            for sname, dff in all_sheets.items():
                                safe = str(sname)[:31]
                                dff.to_excel(w, sheet_name=safe, index=False)
                        work = temp_path
                        self.mapping.append((orig, work))
                        self.temp_files.append(temp_path)
                    except Exception as e2:
                        QMessageBox.warning(self, "Ошибка конвертации", f"Не удалось конвертировать {orig}:\n{e}\n{e2}")
                finally:
                    self.progress.setValue(int(idx / total * 30))
                    QtWidgets.QApplication.processEvents()
            else:
                # already xlsx or cannot use xlwings -> use original
                self.mapping.append((orig, orig))
                self.progress.setValue(int(idx / total * 30))
                QtWidgets.QApplication.processEvents()

        # close app
        try:
            if app is not None:
                app.quit()
        except:
            pass

    # ---------- merge: separate sheets ----------
    def _merge_each_sheet_separate(self, mapping):
        # mapping: list of tuples (original, work)
        total = len(mapping)
        writer = pd.ExcelWriter(self.output_file, engine="openpyxl")
        sheet_index = 1

        for i, (original, work) in enumerate(mapping, 1):
            try:
                xls = pd.ExcelFile(work, engine="openpyxl")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка чтения", f"Не удалось прочитать {work}:\n{e}")
                continue

            base_name = os.path.splitext(os.path.basename(original))[0]

            for sheet in xls.sheet_names:
                try:
                    df = xls.parse(sheet)
                except Exception:
                    continue

                # header with full original filename
                header_df = pd.DataFrame([[f"Полное имя файла: {os.path.basename(original)}"]])
                empty_df = pd.DataFrame([[]])
                final = pd.concat([header_df, empty_df, df], ignore_index=True)

                # safe sheet name (unique)
                safe_file = base_name[:20]
                safe_sheet = str(sheet)[:20]
                candidate = f"{sheet_index}_{safe_file}_{safe_sheet}"
                candidate = candidate[:31]
                # ensure uniqueness within workbook
                while candidate in writer.book.sheetnames:
                    sheet_index += 1
                    candidate = f"{sheet_index}_{safe_file}_{safe_sheet}"[:31]

                final.to_excel(writer, sheet_name=candidate, index=False, header=False)
                sheet_index += 1

            # update progress (30..85)
            self.progress.setValue(int(i / total * 55 + 30))
            QtWidgets.QApplication.processEvents()

        writer.close()

    # ---------- merge: one table ----------
    def _merge_into_one_table(self, mapping):
        total = len(mapping)
        blocks = []
        counter = 1

        for i, (original, work) in enumerate(mapping, 1):
            try:
                xls = pd.ExcelFile(work, engine="openpyxl")
            except Exception as e:
                QMessageBox.warning(self, "Ошибка чтения", f"Не удалось прочитать {work}:\n{e}")
                continue

            for sheet in xls.sheet_names:
                try:
                    df = xls.parse(sheet)
                except Exception:
                    continue

                # add source columns
                df.insert(0, "Source_Sheet", sheet)
                df.insert(0, "Source_File", os.path.basename(original))

                # if sheet name long -> keep original in a column and shorten Source_Sheet
                if len(sheet) > 20:
                    df.insert(0, "Sheet_Name_Original", sheet)
                    df["Source_Sheet"] = f"Sheet_{counter}"
                    counter += 1

                # build block with header line (original filename) + empty row + df
                header_df = pd.DataFrame([[f"Полное имя файла: {os.path.basename(original)}"]])
                empty_df = pd.DataFrame([[]])
                block = pd.concat([header_df, empty_df, df], ignore_index=True)
                blocks.append(block)

            # update progress (30..85)
            self.progress.setValue(int(i / total * 55 + 30))
            QtWidgets.QApplication.processEvents()

        if blocks:
            final = pd.concat(blocks, ignore_index=True)
            # write without header because blocks already include header rows
            final.to_excel(self.output_file, index=False, header=False)

    # ---------- cleanup ----------
    def _cleanup_temp_files(self):
        for f in list(self.temp_files):
            try:
                if os.path.exists(f):
                    os.remove(f)
            except Exception:
                pass
        self.temp_files = []
        # clear mapping
        self.mapping = []

# ---------- run ----------
def main():
    app = QtWidgets.QApplication(sys.argv)
    w = ExcelMerger()
    w.show()
    sys.exit(app.exec_())

if __name__ == "__main__":
    main()