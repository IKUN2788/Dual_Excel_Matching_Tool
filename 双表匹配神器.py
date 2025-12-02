import sys
import os
import csv
from PyQt5 import QtWidgets, QtCore, QtGui

try:
    import pandas as pd
    HAS_PANDAS = True
except Exception:
    HAS_PANDAS = False


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("双表匹配神器")
        self.resize(1200, 720)
        self.pathA = None
        self.pathB = None
        self.columnsA = []
        self.columnsB = []
        self.rowsA = []
        self.rowsB = []
        self.selectedKeyColsA = []
        self.selectedResultColsB = []
        self.sheetA = None
        self.sheetB = None
        self._build_ui()

    def _build_ui(self):
        central = QtWidgets.QWidget()
        self.setCentralWidget(central)
        vbox = QtWidgets.QVBoxLayout(central)

        top_split = QtWidgets.QSplitter(QtCore.Qt.Horizontal)
        vbox.addWidget(top_split, 4)

        self.leftPanel = QtWidgets.QWidget()
        left_layout = QtWidgets.QVBoxLayout(self.leftPanel)
        left_header = QtWidgets.QHBoxLayout()
        self.btnOpenA = QtWidgets.QPushButton("选择A表")
        self.labelPathA = QtWidgets.QLabel("未选择")
        self.labelPathA.setStyleSheet("color: #888")
        self.labelHintA = QtWidgets.QLabel("支持拖拽文件到本区域")
        self.labelHintA.setStyleSheet("color: #888")
        left_header.addWidget(self.btnOpenA)
        left_header.addWidget(self.labelPathA, 1)
        left_header.addWidget(self.labelHintA)
        left_layout.addLayout(left_header)
        self.tableA = QtWidgets.QTableWidget()
        self.tableA.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tableA.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectColumns)
        self.tableA.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.tableA.horizontalHeader().sectionClicked.connect(lambda idx: self._toggle_col('A', idx))
        left_layout.addWidget(self.tableA, 1)
        self.labelSelA = QtWidgets.QLabel("键列：无")
        left_layout.addWidget(self.labelSelA)

        self.rightPanel = QtWidgets.QWidget()
        right_layout = QtWidgets.QVBoxLayout(self.rightPanel)
        right_header = QtWidgets.QHBoxLayout()
        self.btnOpenB = QtWidgets.QPushButton("选择B表")
        self.labelPathB = QtWidgets.QLabel("未选择")
        self.labelPathB.setStyleSheet("color: #888")
        self.labelHintB = QtWidgets.QLabel("支持拖拽文件到本区域")
        self.labelHintB.setStyleSheet("color: #888")
        right_header.addWidget(self.btnOpenB)
        right_header.addWidget(self.labelPathB, 1)
        right_header.addWidget(self.labelHintB)
        right_layout.addLayout(right_header)
        self.tableB = QtWidgets.QTableWidget()
        self.tableB.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        self.tableB.setSelectionBehavior(QtWidgets.QAbstractItemView.SelectColumns)
        self.tableB.setSelectionMode(QtWidgets.QAbstractItemView.NoSelection)
        self.tableB.horizontalHeader().sectionClicked.connect(lambda idx: self._toggle_col('B', idx))
        right_layout.addWidget(self.tableB, 1)
        self.labelSelB = QtWidgets.QLabel("结果列：无")
        right_layout.addWidget(self.labelSelB)

        top_split.addWidget(self.leftPanel)
        top_split.addWidget(self.rightPanel)
        top_split.setSizes([600, 600])

        bottom_panel = QtWidgets.QWidget()
        bottom_layout = QtWidgets.QVBoxLayout(bottom_panel)
        bottom_header = QtWidgets.QHBoxLayout()
        self.labelStatus = QtWidgets.QLabel("匹配状态：等待选择键列")
        bottom_header.addWidget(self.labelStatus, 1)
        self.btnExportCSV = QtWidgets.QPushButton("导出CSV")
        self.btnExportExcel = QtWidgets.QPushButton("导出Excel")
        bottom_header.addWidget(self.btnExportCSV)
        bottom_header.addWidget(self.btnExportExcel)
        bottom_layout.addLayout(bottom_header)
        self.previewTable = QtWidgets.QTableWidget()
        self.previewTable.setEditTriggers(QtWidgets.QAbstractItemView.NoEditTriggers)
        bottom_layout.addWidget(self.previewTable, 2)
        vbox.addWidget(bottom_panel, 3)

        self.btnOpenA.clicked.connect(lambda: self._open_file_dialog('A'))
        self.btnOpenB.clicked.connect(lambda: self._open_file_dialog('B'))
        self.btnExportCSV.clicked.connect(self._export_csv)
        self.btnExportExcel.clicked.connect(self._export_excel)
        self.leftPanel.setAcceptDrops(True)
        self.rightPanel.setAcceptDrops(True)
        self.leftPanel.installEventFilter(self)
        self.rightPanel.installEventFilter(self)

    def eventFilter(self, obj, event):
        if event.type() == QtCore.QEvent.DragEnter:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                return True
        if event.type() == QtCore.QEvent.Drop:
            urls = event.mimeData().urls()
            if urls:
                path = urls[0].toLocalFile()
                sheet = self._choose_sheet(path)
                if obj is self.leftPanel:
                    self._load_file('A', path, sheet)
                elif obj is self.rightPanel:
                    self._load_file('B', path, sheet)
                event.acceptProposedAction()
                return True
        return super().eventFilter(obj, event)

    def _open_file_dialog(self, kind):
        dlg = QtWidgets.QFileDialog(self, "选择文件")
        dlg.setFileMode(QtWidgets.QFileDialog.ExistingFile)
        dlg.setNameFilters(["CSV 文件 (*.csv)", "Excel 文件 (*.xlsx *.xls)", "所有文件 (*.*)"])
        if dlg.exec_():
            paths = dlg.selectedFiles()
            if paths:
                path = paths[0]
                sheet = self._choose_sheet(path)
                self._load_file(kind, path, sheet)

    def _choose_sheet(self, path):
        ext = os.path.splitext(path)[1].lower()
        if ext in ['.xlsx', '.xls'] and HAS_PANDAS:
            try:
                xls = pd.ExcelFile(path)
                names = list(xls.sheet_names)
                if not names:
                    return None
                if len(names) == 1:
                    return names[0]
                item, ok = QtWidgets.QInputDialog.getItem(self, "选择Sheet", "请选择工作表：", names, 0, False)
                return item if ok else names[0]
            except Exception:
                return None
        return None

    def _load_file(self, kind, path, sheet=None):
        if not os.path.isfile(path):
            QtWidgets.QMessageBox.warning(self, "错误", "文件不存在")
            return
        ok, columns, rows, err = self._read_table(path, sheet)
        if not ok:
            QtWidgets.QMessageBox.warning(self, "读取失败", err)
            return
        if kind == 'A':
            self.pathA = path
            self.sheetA = sheet
            self.columnsA = columns
            self.rowsA = rows
            self.labelPathA.setText(path + (f" - {sheet}" if sheet else ""))
            self.selectedKeyColsA = []
            self._render_table(self.tableA, self.columnsA, self.rowsA, 10)
            self.labelSelA.setText("键列：无")
        else:
            self.pathB = path
            self.sheetB = sheet
            self.columnsB = columns
            self.rowsB = rows
            self.labelPathB.setText(path + (f" - {sheet}" if sheet else ""))
            self.selectedResultColsB = []
            self._render_table(self.tableB, self.columnsB, self.rowsB, 10)
            self.labelSelB.setText("结果列：无")
        self._update_status_and_preview()

    def _render_table(self, table, columns, rows, limit):
        table.clear()
        table.setColumnCount(len(columns))
        table.setHorizontalHeaderLabels([str(c) for c in columns])
        n = min(limit, len(rows))
        table.setRowCount(n)
        for i in range(n):
            row = rows[i]
            for j, c in enumerate(columns):
                v = row.get(c, "")
                item = QtWidgets.QTableWidgetItem(str(v))
                table.setItem(i, j, item)
        table.resizeColumnsToContents()

    def _toggle_col(self, kind, idx):
        if kind == 'A':
            if idx < 0 or idx >= len(self.columnsA):
                return
            name = self.columnsA[idx]
            if name in self.selectedKeyColsA:
                self.selectedKeyColsA.remove(name)
                self._highlight_column(self.tableA, idx, False)
            else:
                self.selectedKeyColsA.append(name)
                self._highlight_column(self.tableA, idx, True)
            if self.selectedKeyColsA:
                self.labelSelA.setText("键列：" + ", ".join(self.selectedKeyColsA))
            else:
                self.labelSelA.setText("键列：无")
        else:
            if idx < 0 or idx >= len(self.columnsB):
                return
            name = self.columnsB[idx]
            if name in self.selectedResultColsB:
                self.selectedResultColsB.remove(name)
                self._highlight_column(self.tableB, idx, False)
            else:
                self.selectedResultColsB.append(name)
                self._highlight_column(self.tableB, idx, True)
            if self.selectedResultColsB:
                self.labelSelB.setText("结果列：" + ", ".join(self.selectedResultColsB))
            else:
                self.labelSelB.setText("结果列：无")
        self._update_status_and_preview()

    def _highlight_column(self, table, idx, selected):
        color = QtGui.QColor(255, 230, 180) if selected else QtGui.QColor(255, 255, 255)
        rows = table.rowCount()
        for r in range(rows):
            item = table.item(r, idx)
            if item is None:
                item = QtWidgets.QTableWidgetItem("")
                table.setItem(r, idx, item)
            item.setBackground(color)

    def _update_status_and_preview(self):
        if not self.rowsA or not self.rowsB or not self.selectedKeyColsA:
            self.labelStatus.setText("匹配状态：等待选择键列")
            self.previewTable.clear()
            self._update_export_buttons(False)
            return
        missing = [c for c in self.selectedKeyColsA if c not in self.columnsB]
        if missing:
            self.labelStatus.setText("匹配状态：B表缺少键列：" + ", ".join(missing))
            self.previewTable.clear()
            self._update_export_buttons(False)
            return
        ka = self._key_set(self.rowsA, self.selectedKeyColsA)
        kb = self._key_set(self.rowsB, self.selectedKeyColsA)
        inter = ka & kb
        if len(inter) == len(ka):
            status = "完全包含"
        elif len(inter) > 0:
            status = "部分包含"
        else:
            status = "不包含"
        self.labelStatus.setText(f"匹配状态：{status}（A唯一键{len(ka)}，B包含{len(inter)}）")
        self._update_export_buttons(True)
        self._render_preview()

    def _update_export_buttons(self, enabled):
        if hasattr(self, 'btnExportCSV'):
            self.btnExportCSV.setEnabled(bool(enabled))
        if hasattr(self, 'btnExportExcel'):
            self.btnExportExcel.setEnabled(bool(enabled))

    def _render_preview(self):
        header, res_rows = self._build_result()
        self.previewTable.clear()
        self.previewTable.setColumnCount(len(header))
        self.previewTable.setHorizontalHeaderLabels(header)
        limit = min(200, len(res_rows))
        self.previewTable.setRowCount(limit)
        for i in range(limit):
            row = res_rows[i]
            for j, v in enumerate(row):
                item = QtWidgets.QTableWidgetItem(str(v))
                self.previewTable.setItem(i, j, item)
        self.previewTable.resizeColumnsToContents()

    def _build_result(self):
        if not self.rowsA or not self.rowsB or not self.selectedKeyColsA:
            return [], []
        if any(c not in self.columnsB for c in self.selectedKeyColsA):
            return [], []
        b_map = {}
        for row in self.rowsB:
            k = tuple(str(row.get(c, "")) for c in self.selectedKeyColsA)
            b_map.setdefault(k, []).append(row)
        res_rows = []
        a_cols = self.columnsA
        b_cols = [c for c in self.selectedResultColsB]
        for a in self.rowsA:
            k = tuple(str(a.get(c, "")) for c in self.selectedKeyColsA)
            if k in b_map:
                for b in b_map[k]:
                    res_rows.append([str(a.get(c, "")) for c in a_cols] + [str(b.get(c, "")) for c in b_cols])
        header = [str(c) for c in a_cols] + ["B:" + str(c) for c in b_cols]
        return header, res_rows

    def _export_csv(self):
        header, rows = self._build_result()
        if not header:
            QtWidgets.QMessageBox.warning(self, "导出失败", "请先选择有效的键列并确保B表包含这些列")
            return
        base_dir = os.path.dirname(self.pathA or self.pathB or os.getcwd())
        default = os.path.join(base_dir, "匹配结果.csv")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出CSV", default, "CSV 文件 (*.csv)")
        if not path:
            return
        try:
            with open(path, "w", encoding="utf-8-sig", newline="") as f:
                w = csv.writer(f)
                w.writerow(header)
                for r in rows:
                    w.writerow(r)
            QtWidgets.QMessageBox.information(self, "导出成功", path)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "导出失败", str(e))

    def _export_excel(self):
        header, rows = self._build_result()
        if not header:
            QtWidgets.QMessageBox.warning(self, "导出失败", "请先选择有效的键列并确保B表包含这些列")
            return
        base_dir = os.path.dirname(self.pathA or self.pathB or os.getcwd())
        default = os.path.join(base_dir, "匹配结果.xlsx")
        path, _ = QtWidgets.QFileDialog.getSaveFileName(self, "导出Excel", default, "Excel 文件 (*.xlsx)")
        if not path:
            return
        if HAS_PANDAS:
            try:
                import pandas as pd
                df = pd.DataFrame(rows, columns=header)
                df.to_excel(path, index=False)
                QtWidgets.QMessageBox.information(self, "导出成功", path)
                return
            except Exception as e:
                pass
        try:
            from openpyxl import Workbook
            wb = Workbook()
            ws = wb.active
            ws.title = "结果"
            ws.append(header)
            for r in rows:
                ws.append(r)
            wb.save(path)
            QtWidgets.QMessageBox.information(self, "导出成功", path)
        except Exception as e:
            QtWidgets.QMessageBox.warning(self, "导出失败", str(e))

    def _key_set(self, rows, cols):
        s = set()
        for r in rows:
            s.add(tuple(str(r.get(c, "")) for c in cols))
        return s

    def _read_table(self, path, sheet_name=None):
        ext = os.path.splitext(path)[1].lower()
        if ext in ['.xlsx', '.xls']:
            if not HAS_PANDAS:
                return False, [], [], "读取Excel需要pandas，请安装后重试"
            try:
                df = pd.read_excel(path, dtype=str, sheet_name=sheet_name)
                df = df.fillna('')
                columns = [str(c) for c in df.columns]
                rows = [dict((str(k), str(v)) for k, v in row.items()) for row in df.to_dict(orient='records')]
                return True, columns, rows, None
            except Exception as e:
                return False, [], [], str(e)
        elif ext in ['.csv', '.txt']:
            encodings = ['utf-8', 'utf-8-sig', 'gb18030']
            last_err = None
            for enc in encodings:
                try:
                    with open(path, 'r', encoding=enc, newline='') as f:
                        sample = f.read(1024)
                        f.seek(0)
                        try:
                            dialect = csv.Sniffer().sniff(sample)
                        except Exception:
                            dialect = csv.excel
                        reader = csv.DictReader(f, dialect=dialect)
                        columns = [str(c) for c in reader.fieldnames] if reader.fieldnames else []
                        rows = []
                        for row in reader:
                            r = {}
                            for c in columns:
                                v = row.get(c, '')
                                r[str(c)] = '' if v is None else str(v)
                            rows.append(r)
                        return True, columns, rows, None
                except Exception as e:
                    last_err = e
                    continue
            return False, [], [], str(last_err) if last_err else "未知错误"
        else:
            if HAS_PANDAS:
                try:
                    df = pd.read_csv(path, dtype=str, encoding='utf-8')
                    df = df.fillna('')
                    columns = [str(c) for c in df.columns]
                    rows = [dict((str(k), str(v)) for k, v in row.items()) for row in df.to_dict(orient='records')]
                    return True, columns, rows, None
                except Exception as e:
                    return False, [], [], str(e)
            return False, [], [], "不支持的文件类型"


def main():
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    w.show()
    sys.exit(app.exec_())


if __name__ == '__main__':
    main()

