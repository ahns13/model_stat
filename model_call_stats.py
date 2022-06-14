from PyQt5 import QtWidgets
from PyQt5 import uic, QtCore, QtGui
from PyQt5.QtWidgets import *
import sys, os
import openpyxl as xl
import re
from model_stat_chart import MyChart

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

form_class = uic.loadUiType(BASE_DIR + r'\model_call_stats_main.ui')[0]


class TableDelegate(QtWidgets.QStyledItemDelegate):
    def initStyleOption(self, option, index):
        super(TableDelegate, self).initStyleOption(option, index)
        option.font.setPixelSize(12)


class MainWindow(QMainWindow, form_class):
    def __init__(self):
        super(MainWindow, self).__init__()
        self.setupUi(self)

        self.excel_info = {
            "start_row": 4,
            "start_col": 1,
            "end_col": 14  # 비고
        }

        self.tableWidget.setColumnWidth(0, 100)
        self.tableWidget.setColumnWidth(1, 45)
        self.tableWidget.setColumnWidth(2, 65)
        self.tableWidget.setColumnWidth(3, 65)
        self.tableWidget.setColumnWidth(4, 130)
        self.tableWidget.setColumnWidth(5, 80)
        self.tableWidget.setColumnWidth(6, 70)
        self.tableWidget.setColumnWidth(7, 120)
        self.tableWidget.setColumnWidth(8, 120)
        self.tableWidget.setColumnWidth(9, 110)
        self.tableWidget.setColumnWidth(10, 60)
        self.tableWidget.setColumnWidth(11, 60)
        self.tableWidget.setColumnWidth(12, 120)
        self.tableWidget.horizontalHeader().setFont(QtGui.QFont("", 8))
        self.tableWidget.verticalHeader().setDefaultSectionSize(15)
        self.tableWidget.verticalHeader().setMinimumSectionSize(22)
        self.tableWidget.setWordWrap(False)

        table_delegate = TableDelegate(self.tableWidget)
        self.tableWidget.setItemDelegate(table_delegate)

        self.btn_file.clicked.connect(self.fileSelector)
        self.btn_init.clicked.connect(self.filterInit)
        self.tableData = []
        self.tableFilteredValue = {idx: "" for idx in range(self.tableWidget.columnCount())}
        self.tableFilteredStackNo = 0

        self.tableFilterData = {
            0: [],  # 날짜
            1: ["월", "화", "수", "목", "금", "토", "일"],
            2: ["오전", "오후"],
            3: [],  # 접근 경로
            4: "",
            5: [],  # 섭외 종류
            6: [],  # 업체 성격
            7: "", 8: "", 9: "",
            10: ["O", "X", "빈값"],
            11: ["O", "X", "빈값"],
            12: ""

        }
        self.filter_exec_list = {}
        for key in self.tableFilterData.keys():
            key_data = self.tableFilterData[key]
            if type(key_data) == list:
                self.filter_exec_list[key] = False if len(key_data) else True
            else:
                self.filter_exec_list[key] = None

        # chart
        self.btn_monthly_report.clicked.connect(lambda: self.chartDialog())
        self.chart_col_map = {
            "요일": 1, "오전/오후": 2, "접근경로": 3, "섭외종류": 5, "진행여부": 10, "최종완료": 11
        }
        self.comboBox_cols.addItems(self.chart_col_map.keys())
        self.btn_monthly_item_report.clicked.connect(lambda p_change_value: self.chartDialog(self.comboBox_cols.currentText()))

        self.show()

    def fileSelector(self):
        select_file = QFileDialog.getOpenFileName(None, "키워드 광고 파일 열기", "", "Excel Files (*.xlsx *.xls)")
        self.label_fileName.setText(os.path.basename(select_file[0]))

        wb = xl.load_workbook(select_file[0])
        for sh_idx, sheet in enumerate(wb.worksheets):
            if re.search(re.compile(r"[0-9]{0,2}월"), sheet.title) is not None:
                for r in list(sheet.rows)[self.excel_info["start_row"]:]:
                    row_data = []
                    if r[self.excel_info["start_col"]].value is not None:
                        for c_idx, c in enumerate(r[self.excel_info["start_col"]:self.excel_info["end_col"]]):
                            if c_idx == 0:  # 날짜
                                date_value = c.value
                                date_month = str(date_value.month)+"월"
                                date_text = str(date_value.year) + "년 " + date_month + str(date_value.day) + "일"
                                row_data.append(date_text)

                                if c_idx in self.tableFilterData.keys():
                                    if date_month not in self.tableFilterData[c_idx]:
                                        self.tableFilterData[c_idx].append(date_month)
                            else:
                                row_data.append(c.value if c.value is not None else "")
                                if c_idx in self.tableFilterData.keys():
                                    if self.filter_exec_list[c_idx]:
                                        c.value = c.value if c.value else "빈값"
                                        if c.value not in self.tableFilterData[c_idx]:
                                            self.tableFilterData[c_idx].append(c.value)

                        self.tableData.append(row_data)

        self.label_total_count.setText(self.label_total_count.text() + str(len(self.tableData)))
        wb.close()

        for key in self.tableFilterData.keys():
            if self.filter_exec_list[key]:
                self.tableFilterData[key].sort()
            if self.filter_exec_list[key] is not None:
                self.tableFilterData[key].insert(0, "전체")

        for r_idx, r in enumerate(self.tableData):
            self.tableWidget.insertRow(r_idx+1)
            for c_idx, c in enumerate(range(self.tableWidget.columnCount())):
                self.tableWidget.setItem(r_idx+1, c_idx, QTableWidgetItem(str(r[c_idx])))

        for c_idx in range(self.tableWidget.columnCount()):
            key_data = self.filter_exec_list[c_idx]
            if key_data is not None:
                cellWidget = QComboBox()
                cellWidget.addItems(self.tableFilterData[c_idx])
                colWidth = self.tableWidget.columnWidth(c_idx)
                cellWidget.setStyleSheet("""
                        QComboBox QAbstractItemView { min-width: """ + str(colWidth) + """px; }
                """)
                cellWidget.currentTextChanged.connect(
                    lambda p_change_value, p_col_idx=c_idx: self.filterExec(p_col_idx, p_change_value)
                )
            else:
                cellWidget = QLineEdit()
                cellWidget.returnPressed.connect(
                    lambda p_col_idx=c_idx, p_line_obj=cellWidget: self.searchExecLE(p_col_idx, p_line_obj)
                )
            self.tableWidget.setCellWidget(0, c_idx, cellWidget)

    def filterExec(self, v_table_col_index, v_filter_value):
        if v_filter_value == "전체":
            if self.tableFilteredStackNo:
                self.tableFilteredStackNo -= 1
                self.tableFilteredValue[v_table_col_index] = ""
                for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
                    if self.tableWidget.rowHeight(r_idx) == 0:
                        show_row = True
                        for key in self.tableFilteredValue.keys():
                            filter_value = self.tableFilteredValue[key]
                            if self.filter_exec_list[key] is not None:
                                if filter_value and self.tableWidget.item(r_idx, key).text() != filter_value:
                                    show_row = False
                            else:
                                if filter_value and filter_value not in self.tableWidget.item(r_idx, key).text():
                                    show_row = False
                            if not show_row:
                                break
                        if show_row:
                            self.tableWidget.showRow(r_idx)
        else:
            self.tableFilteredStackNo += 1
            self.tableFilteredValue[v_table_col_index] = v_filter_value
            for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
                cell_text = self.tableWidget.item(r_idx, v_table_col_index).text()
                cell_text = cell_text if cell_text else "빈값"
                if v_table_col_index == 0:  # date
                    if v_filter_value not in cell_text:
                        self.tableWidget.hideRow(r_idx)
                elif cell_text != v_filter_value:
                    self.tableWidget.hideRow(r_idx)
                elif cell_text == v_filter_value:
                    self.tableWidget.showRow(r_idx)

    def searchExecLE(self, v_table_col_index, v_line_obj):
        input_text = v_line_obj.text()
        if input_text:
            self.tableFilteredValue[v_table_col_index] = input_text
            self.tableFilteredStackNo += 1
            for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
                if input_text not in self.tableWidget.item(r_idx, v_table_col_index).text():
                    self.tableWidget.hideRow(r_idx)
        else:
            self.filterExec(v_table_col_index, "전체")

    def filterInit(self):
        if self.label_fileName.text():
            for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
                if self.tableWidget.rowHeight(r_idx) == 0:
                    self.tableWidget.showRow(r_idx)

            for idx in range(self.tableWidget.columnCount()):
                self.tableFilteredValue[idx] = ""
                if self.filter_exec_list[idx] is not None:
                    self.tableWidget.cellWidget(0, idx).setCurrentIndex(0)
                else:
                    self.tableWidget.cellWidget(0, idx).setText("")

    def chartDialog(self, v_select_col=None):
        chart_data = {
            "filter_count": self.tableWidget.rowCount(),
            "monthly": {}
        }

        monthly_data = []
        for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
            if self.tableWidget.rowHeight(r_idx) > 0:
                monthly_data.append(self.tableWidget.item(r_idx, 0).text())

        for data in monthly_data:
            month = re.search(re.compile(r"[0-9]{0,2}월"), data).group()
            if month in chart_data["monthly"]:
                chart_data["monthly"][month] += 1
            else:
                chart_data["monthly"][month] = 0

        if v_select_col is not None:
            selecet_data = []
            distinct_data = []
            chart_data[v_select_col] = {}
            col_index = self.chart_col_map[v_select_col]
            for r_idx in list(range(self.tableWidget.rowCount()))[1:]:
                if self.tableWidget.rowHeight(r_idx) > 0:
                    r_value = self.tableWidget.item(r_idx, col_index).text()
                    r_value = r_value if r_value else "빈값"
                    selecet_data.append(r_value)
                    if r_value not in distinct_data:
                        distinct_data.append(r_value)

            distinct_data = sorted(distinct_data, key=lambda x: self.tableFilterData[col_index].index(x))

            for d in distinct_data:
                chart_data[v_select_col][d] = {}
            col_data = chart_data[v_select_col]

            for key in col_data.keys():
                col_data[key] = {m: 0 for m in chart_data["monthly"].keys()}
            for idx, data in enumerate(selecet_data):
                month = re.search(re.compile(r"[0-9]{0,2}월"), monthly_data[idx]).group()
                col_data[data][month] += 1

        chart = MyChart(chart_data, v_select_col)
        chart.exec_()


if __name__ == "__main__":
    app = QtWidgets.QApplication(sys.argv)
    main_window = MainWindow()
    app.exec_()