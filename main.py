import openpyxl
import sys
from PyQt5.QtWidgets import QApplication, QMainWindow, QPushButton, QDialog, QFileDialog, QVBoxLayout, QWidget


class MyWidget(QWidget):
    def __init__(self):
        super().__init__()
        self.resize(300,650)

        # Создаем 13 кнопок
        self.buttons = []
        for month in ["Январь", "Февраль", "Март", "Апрель", "Май", "Июнь", "Июль", "Август", "Сентябрь", "Октябрь",
                      "Ноябрь", "Декабрь", "Год"]:
            button = QPushButton(f"Выберите {month}", self)
            self.buttons.append(button)

        # Создаем 13 переменных для хранения путей к файлам
        self.jan = ""
        self.feb = ""
        self.mar = ""
        self.apr = ""
        self.may = ""
        self.jun = ""
        self.jul = ""
        self.aug = ""
        self.sep = ""
        self.oct = ""
        self.nov = ""
        self.dec = ""
        self.god = ""

        self.transport_button = QPushButton("Выполнить", self)
        self.transport_button.move(100, 618)
        self.transport_button.clicked.connect(self.excel_transport)

        # Создаем вертикальный layout и добавляем в него кнопки
        layout = QVBoxLayout(self)
        for button in self.buttons:
            layout.addWidget(button)

        # Подключаем слоты к сигналам кнопок
        for i, button in enumerate(self.buttons):
            button.clicked.connect(lambda _, i=i: self.get_file(i))

    def get_file(self, index):
        # Открываем диалоговое окно для выбора файла
        file_dialog = QFileDialog()
        file_dialog.setFileMode(QFileDialog.ExistingFile)
        file_dialog.setNameFilter("Excel Files (*.xlsx)")
        if file_dialog.exec_():
            # Получаем имя выбранного файла
            selected_file = file_dialog.selectedFiles()[0]
            selected_file = selected_file.replace("/", "\\")

            # Записываем путь к файлу в отдельную переменную
            if index == 0:
                self.jan = selected_file
            elif index == 1:
                self.feb = selected_file
            elif index == 2:
                self.mar = selected_file
            elif index == 3:
                self.apr = selected_file
            elif index == 4:
                self.may = selected_file
            elif index == 5:
                self.jun = selected_file
            elif index == 6:
                self.jul = selected_file
            elif index == 7:
                self.aug = selected_file
            elif index == 8:
                self.sep = selected_file
            elif index == 9:
                self.oct = selected_file
            elif index == 10:
                self.nov = selected_file
            elif index == 11:
                self.dec = selected_file
            elif index == 12:
                self.god = selected_file

            print(f"Путь к файлу для {self.buttons[index].text()}:")
            print(selected_file)

            # Создаем кнопку для запуска функции excel_transport()

    def excel_transport(self):
        book_january = openpyxl.open(f"{self.jan}", data_only=True, read_only=True)
        book_february = openpyxl.open(f"{self.feb}", data_only=True, read_only=True)
        book_march = openpyxl.open(f"{self.mar}", data_only=True, read_only=True)
        book_april = openpyxl.open(f"{self.apr}", data_only=True, read_only=True)
        book_may = openpyxl.open(f"{self.may}", data_only=True, read_only=True)
        book_june = openpyxl.open(f"{self.jun}", data_only=True, read_only=True)
        book_july = openpyxl.open(f"{self.jul}", data_only=True, read_only=True)
        book_august = openpyxl.open(f"{self.aug}", data_only=True, read_only=True)
        book_september = openpyxl.open(f"{self.sep}", data_only=True, read_only=True)
        book_october = openpyxl.open(f"{self.oct}", data_only=True, read_only=True)
        book_november = openpyxl.open(f"{self.nov}", data_only=True, read_only=True)
        book_december = openpyxl.open(f"{self.dec}", data_only=True, read_only=True)
        book_years = openpyxl.load_workbook(f"{self.god}")

        sheet_years = book_years.active
        sheet_january = book_january.active
        sheet_february = book_february.active
        sheet_march = book_march.active
        sheet_april = book_april.active
        sheet_may = book_may.active
        sheet_june = book_june.active
        sheet_july = book_july.active
        sheet_august = book_august.active
        sheet_september = book_september.active
        sheet_october = book_october.active
        sheet_november = book_november.active
        sheet_december = book_december.active

        i1 = 1
        i = 5
        while i1 <= 14:
            i = 5
            if i1 == 1:
                s = "F"
                m = sheet_january['E5':'E90']
            if i1 == 2:
                s = "G"
                m = sheet_february['E5':'E90']
                print("february")
            if i1 == 3:
                s = "H"
                m = sheet_march['E5':'E90']
                print("march")
            if i1 == 5:
                s = "J"
                m = sheet_april['E5':'E90']
                print("april")
            if i1 == 6:
                s = "K"
                m = sheet_may['E5':'E90']
                print("may")
            if i1 == 7:
                s = "L"
                m = sheet_june['E5':'E90']
                print("june")
            if i1 == 8:
                s = "N"
                m = sheet_july['E5':'E90']
                print("july")
            if i1 == 9:
                s = "O"
                m = sheet_august['E5':'E90']
                print("august")
            if i1 == 10:
                s = "P"
                m = sheet_september['E5':'E90']
                print("september")
            if i1 == 12:
                s = "R"
                m = sheet_october['E5':'E90']
                print("october")
            if i1 == 13:
                s = "S"
                m = sheet_november['E5':'E90']
                print("november")
            if i1 == 14:
                s = "T"
                m = sheet_december['E5':'E90']
                print("december")
            cells = m
            for cell in cells:
                if cell[0].coordinate == f'E15':  # пропускаем 15 ячейку
                    continue
                print(cell[0].value)
                sheet_years[f'{s}{i}'] = cell[
                    0].value  # если i1 = 1 то F | если i1 = 2 то G | если i1 = 3 то H | если i1 = 5 то J | усли i1 = 6 то K | если i1 = 7 то L
                i += 1  # если i1 = 8 то N | если i1 = 9 то O | усли i1 = 10 то P | усли i1 = 12 то R | если i1 = 13 то S | если i1 = 14 то T
            i1 += 1
        book_years.save('год.xlsx')
        print("Функция excel_transport() запущена")


if __name__ == "__main__":
    app = QApplication([])
    widget = MyWidget()
    widget.show()
    app.exec_()