from PyQt5.QtWidgets import *
from PyQt5 import QtWidgets
from PyQt5 import QtCore
from PyQt5.QtChart import *
from PyQt5.QtGui import QPainter
import sys


class MyChart(QDialog):
    def __init__(self, v_chart_data, v_select_col=None):
        super(MyChart, self).__init__()

        # window size
        self.setMinimumSize(800, 500)

        month_list = v_chart_data["monthly"].keys()

        # total_chart
        if v_select_col is None:
            total_data = sum([v_chart_data["monthly"][m] for m in month_list])
        else:
            total_data = []
            for key in v_chart_data[v_select_col].keys():
                total_data.append(sum([v_chart_data[v_select_col][key][m] for m in month_list]))
        set_total = QBarSet("")
        set_total.append(total_data)
        total_series = QBarSeries()
        total_series.append(set_total)
        total_series.setLabelsVisible(True)
        total_series.setLabelsPosition(1)

        total_chart = QChart()
        total_chart.legend().hide()
        total_chart.addSeries(total_series)

        total_xAxis = QBarCategoryAxis()
        if v_select_col is None:
            total_xAxis.append("전체")
        else:
            total_xAxis.append(v_chart_data[v_select_col].keys())
        total_chart.addAxis(total_xAxis, QtCore.Qt.AlignBottom)
        total_series.attachAxis(total_xAxis)

        total_chart_view = QChartView(total_chart)
        total_chart_view.setRenderHint(QPainter.Antialiasing)

        # monthly_chart
        if v_select_col is None:
            raw_data = [v_chart_data["monthly"][key] for key in month_list]
            set0 = QBarSet('건 수')
            set0.append(raw_data)
            series = QBarSeries()
            series.append(set0)
        else:
            set_cols = []
            for key in v_chart_data[v_select_col].keys():
                item_set = QBarSet(key)
                item_set.append([v_chart_data[v_select_col][key][m] for m in month_list])
                set_cols.append(item_set)
            series = QBarSeries()
            for set in set_cols:
                series.append(set)
        series.setLabelsVisible(True)
        series.setLabelsPosition(1)

        chart = QChart()
        chart.addSeries(series)

        xAxis = QBarCategoryAxis()
        xAxis.append(month_list)
        chart.addAxis(xAxis, QtCore.Qt.AlignBottom)
        series.attachAxis(xAxis)

        yAxis = QValueAxis()
        chart.addAxis(yAxis, QtCore.Qt.AlignLeft)
        series.attachAxis(yAxis)

        # displaying chart
        chart_view = QChartView(chart)
        chart_view.setRenderHint(QPainter.Antialiasing)

        self.layout = QHBoxLayout()
        self.layout.addWidget(total_chart_view, 1)
        self.layout.addWidget(chart_view, 1)
        self.setLayout(self.layout)
        self.show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    win = MyChart()
    app.exec_()