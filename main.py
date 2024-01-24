import io
import sys
from tkinter.filedialog import askopenfilename
import numpy as np
import openpyxl
from PySide6 import QtGui
from PySide6.QtCore import Qt
from PySide6.QtGui import QColor
from PySide6.QtWidgets import QMainWindow, QApplication, QLabel, QVBoxLayout, QTableWidget, QWidget, QTableWidgetItem

from class_cable import Cable
from variables import Variables


class WindowSection(QMainWindow):
    def __init__(self):
        super().__init__()
        self.dict_of_cables = {}
        self.x = []
        self._graph_scale_x: float = 1
        self._graph_scale_y: float = 1
        self._graph_x0: float = 0
        self._graph_y0: float = 0
        self._label = QLabel('hello')
        self._table = QTableWidget()

        self.canvas = QtGui.QPixmap(Variables.screen_bh[0], Variables.screen_bh[1])
        self.canvas.fill(Qt.black)
        self._label.setPixmap(self.canvas)
        self.painter: QtGui.QPainter | None = None
        self.general_layout = QVBoxLayout()
        self.general_layout.addWidget(self._label)
        self.general_layout.addWidget(self._table)
        widget = QWidget()
        widget.setLayout(self.general_layout)
        self.setCentralWidget(widget)
        self.open_file()
        self.init_table()
        self._scale_the_graph()
        self._draw_graph()

    def _draw_graph(self):
        canvas = self._label.pixmap()
        canvas.fill(Qt.black)
        self.painter = QtGui.QPainter(canvas)
        pen = QtGui.QPen()

        for key, cable in self.dict_of_cables.items():
            color = cable.color
            pen.setColor(QtGui.QColor(QColor(*color)))
            if cable.selected:
                pen.setWidth(3)
            else:
                pen.setWidth(1)
            self.painter.setPen(pen)
            self._draw_a_cable(cable)
        self.painter.end()
        self._label.setPixmap(canvas)

    def _draw_a_cable(self, cable: Cable):
        coordinate: [] = cable.coordinate
        #print(f'cable Nr {cable.name}')
        for i in range(len(coordinate) - 1):
            x1 = self._graph_scale_x * (coordinate[i][0] - self._graph_x0) + 10
            x2 = self._graph_scale_x * (coordinate[i + 1][0] - self._graph_x0) + 10
            y1 = Variables.screen_bh[1] - self._graph_scale_y * (coordinate[i][1] - self._graph_y0) - 10
            y2 = Variables.screen_bh[1] - self._graph_scale_y * (coordinate[i + 1][1] - self._graph_y0) - 10
            self.painter.drawLine(x1, y1, x2, y2)
            #print(i, '%0.0f' % x1, '%0.0f' % y1, '%0.0f' % x2, '%0.0f' % y2)

    def _scale_the_graph(self):
        if len(self.x) == 0:
            return None
        self._graph_scale_x = (Variables.screen_bh[0] - 20) / self.x[-1]
        self._graph_x0 = self.x[0]
        y_min = 1000
        y_max = -1000
        for key, cable in self.dict_of_cables.items():
            coordinate: [] = cable.coordinate
            for x, y in coordinate:
                if y > y_max:
                    y_max = y
                elif y < y_min:
                    y_min = y
        self._graph_y0 = y_min
        self._graph_scale_y = (Variables.screen_bh[1] - 20) / (y_max - y_min)

    def open_file(self):
        file_name = askopenfilename()
        print("the file is open -")
        print(file_name)

        with (open(file_name, "rb") as f):
            in_mem_file = io.BytesIO(f.read())
            work_book = openpyxl.load_workbook(in_mem_file, read_only=True)
            print(work_book)
            print(work_book.sheetnames)
            #  for i, sheet in enumerate(work_book.sheetnames):
            self.make_data(sheet=work_book.worksheets[0], sheet_name=work_book.sheetnames[0])
            work_book.close()

    def make_data(self, sheet, sheet_name):
        dict_of_cable = {}
        set_of_current_cables: set[int] = set()
        first_value_x_direction = 4
        i = first_value_x_direction
        x = 0
        y_for_x_value = 3
        while sheet.cell(row=y_for_x_value, column=i).value is not None:
            x += float(sheet.cell(row=y_for_x_value, column=i).value)
            self.x.append(x)
            i += 1

        for y in range(4, 10):
            i = first_value_x_direction
            x = 0
            while sheet.cell(row=y_for_x_value, column=i).value is not None:
                x += float(sheet.cell(row=y_for_x_value, column=i).value)
                value = str(sheet.cell(row=y, column=i).value)
                if value is None or value == 'None':
                    i += 1
                    continue
                if value[-1] == ')':
                    n = value.find(' ')
                    y_i = value[:n]
                    nn = value[n + 1:]
                    y_i = y_i.replace(',', '.')
                    y_i = float(y_i)
                    nn = nn[1: -1]
                    if nn.find(Variables.end) != -1:
                        end_begin = Variables.end
                        nn = nn.replace(Variables.end, '')
                    else:
                        end_begin = Variables.begin
                        nn = nn.replace(Variables.begin, '')
                    numbers_of_cable = nn.split(', ')
                    numbers_of_cable = [int(x) for x in numbers_of_cable]
                    for number_of_cable in set_of_current_cables:
                        cable: Cable = dict_of_cable[number_of_cable]
                        cable.coordinate.append([x, y_i])
                    if 29 in numbers_of_cable:
                        print('i got it')
                    for number_of_cable in numbers_of_cable:
                        if number_of_cable in dict_of_cable:  # remove an old cable
                            cable: Cable = dict_of_cable[number_of_cable]
                            cable.i_n = i - first_value_x_direction
                            set_of_current_cables.discard(number_of_cable)
                        else:  # add a new cable
                            cable = Cable(name=number_of_cable)
                            cable.color = list(np.random.choice(range(256), size=3))
                            dict_of_cable[number_of_cable] = cable
                            cable.i_0 = i - first_value_x_direction
                            set_of_current_cables.add(number_of_cable)
                            cable.coordinate.append([x, y_i])
                        if end_begin == Variables.end:
                            if cable.end:
                                print('!!!cable end', cable.name)
                            cable.end = x
                        elif end_begin == Variables.begin:
                            if cable.begin:
                                print('!!!cable begin', cable.name)
                            cable.begin = x

                else:
                    y_i = float(value)
                    for name_of_cable in set_of_current_cables:
                        cable: Cable = dict_of_cable[name_of_cable]
                        cable.coordinate.append([x, y_i])
                i += 1

        self.dict_of_cables = dict_of_cable
        # for key, cable in self.dict_of_cables.items():
        #    print(cable)

    def init_table(self):
        self._table.setFixedWidth(Variables.screen_bh[0])
        self._table.setColumnCount(len(self.x))
        self._table.setRowCount(len(self.dict_of_cables))
        self._table.itemSelectionChanged.connect(self._selection_changed)
        for i in range(len(self.x)):
            self._table.setColumnWidth(i, 30)
        horizontal_header = ['%0.2f' % x for x in self.x]

        self._table.setHorizontalHeaderLabels(horizontal_header)
        vertical_header = []
        for key in self.dict_of_cables:
            vertical_header.append(str(key))
        self._table.setVerticalHeaderLabels(vertical_header)
        # self._table.itemChanged.connect(self._table_item_is_changed)
        for row, key in enumerate(self.dict_of_cables):
            self._add_a_cable_to_table(cable=self.dict_of_cables[key], row=row)

    def _add_a_cable_to_table(self, cable: Cable, row: int):
        for i in range(cable.i_0, cable.i_n + 1):
            self._table.setItem(row, i, QTableWidgetItem(str(cable.coordinate[i - cable.i_0][1])))

    def _selection_changed(self):
        list_of_selected_items = self._table.selectedIndexes()
        rows = [x.row() for x in list_of_selected_items]
        list_of_hash_ids = list(self.dict_of_cables)
        for key, cable in self.dict_of_cables.items():
            cable.selected = False
        for row in rows:
            cable: Cable = self.dict_of_cables[list_of_hash_ids[row]]
            print(cable.name)
            cable.selected = True

        self._draw_graph()


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = WindowSection()
    window.show()
    app.exec()
