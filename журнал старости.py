import sys
import os
import xlwt
import xlrd
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from PyQt5 import QtWidgets, uic, QtGui
import matplotlib


class MainWindow(QtWidgets.QMainWindow):
    """
    Головне меню програми.
    В ньому є дві кнопки, одна із них дозволяє створити нову групу, друга дозволяє переглядати вже створені.
    """
    def __init__(self, parent=None):
        QtWidgets.QMainWindow.__init__(self, parent)
        self.ui = uic.loadUi('Електронний журнал старости.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        # Натискання на кнопки Створити та Переглянути.
        self.ui.pushButton.setToolTip('<b>Створити новий журнал.</b>')
        self.ui.pushButton_2.setToolTip('<b>Переглянути журнал.</b>')
        self.ui.pushButton.clicked.connect(self.click_1)
        self.ui.pushButton_2.clicked.connect(self.click_2)
        # Показуємо вікно
        self.ui.show()

    def click_1(self):
        """
        Слот, який повертає вікно для створенням групи
        """
        DialogWindow(self)

    def click_2(self):
        """
        Слот, який повертає вікно в якому можна перелянути вже створену групу.
        """
        try:
            WeekChoose(self)
            self.ui.close()
        except FileNotFoundError:
            self.msgWarn = QtWidgets.QMessageBox()
            self.msgWarn.setIcon(QtWidgets.QMessageBox.Warning)
            self.msgWarn.setText("Для перегляду, створіть спочатку групу!")
            self.msgWarn.setStandardButtons(QtWidgets.QMessageBox.Ok)
            self.msgWarn.buttonClicked.connect(self.msgWarn.close)
            self.msgWarn.exec_()


class DialogWindow(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(DialogWindow, self).__init__(parent)
        self.ui = uic.loadUi('Діалог створення журналу.ui')
        self.ui.setWindowTitle('Dialog')
        self.ui.buttonBox.accepted.connect(self.click_1)
        self.ui.buttonBox.rejected.connect(self.click_2)
        self.ui.show()

    def click_1(self):
        GroupData(self)
        # Закриття поточного вікна
        self.parent().ui.close()
        self.ui.close()

    def click_2(self):
        self.ui.close()


class GroupData(QtWidgets.QMainWindow):
    """
    Вікно в якому можна заповнити дані про групу.
    """
    def __init__(self, parent=None):
        super(GroupData, self).__init__(parent)
        self.ui = uic.loadUi('Заповнення даних про групу.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        # Натискання кнопок Створити та Назад.
        self.ui.pushButton.clicked.connect(self.click_1)
        self.ui.pushButton_2.clicked.connect(self.click_2)
        # Показуємо вікно
        self.ui.show()

    def click_1(self):
        """
        Слот, який записує дані з qLineEdit в файл з даними про групу.
        """
        try:
            os.remove('group.txt')
        except FileNotFoundError:
            pass
        # Записуємо назву групи.
        name_of_group = self.ui.lineEdit.text()
        # Записуємо ім'я старости групи.
        name = self.ui.lineEdit_2.text()
        # Записуємо кількість студентів
        number_of_students = self.ui.lineEdit_3.text()
        try:
            assert int(number_of_students) <= 20
        # Записуємо кількість навчальних тижнів.
            number_of_weeks = self.ui.lineEdit_4.text()
            assert int(number_of_weeks) <= 15
            data = [name_of_group, name, number_of_students, number_of_weeks]
            # Створюємо новий файл, якщо його немає, якщо є, то записуємо дані в нього.
            with open('group.txt', 'w') as f:
                for i in data:
                    f.write(i + '\n')
            StudentData(self)
            # Закриття поточного вікна.
            self.ui.close()
        except AssertionError:
            self.msgWarn = QtWidgets.QMessageBox()
            self.msgWarn.setIcon(QtWidgets.QMessageBox.Warning)
            self.msgWarn.setText("Ви ввели не коректні дані.")
            self.msgWarn.setStandardButtons(QtWidgets.QMessageBox.Ok)
            self.msgWarn.buttonClicked.connect(self.msgWarn.close)
            self.msgWarn.exec_()

    def click_2(self):
        MainWindow(self)
        # Закриття поточного вікна.
        self.ui.close()


class StudentData(QtWidgets.QMainWindow):
    """
    Вікно для заповнення таблиці студентами.
    """
    def __init__(self, parent=None):
        super(StudentData, self).__init__(parent)
        self.ui = uic.loadUi('Заповнення студентів.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        # Відкриваємо файл з даними про групу, щоб дізнатись кількість студентів
        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        # Задаємо кількість колонок.
        self.ui.tableWidget.setColumnCount(1)
        # Задаємо кількість рядків, опираючись на кількість студентів в групі.
        self.rows = int(form[2])
        self.weeks = int(form[3])
        self.ui.tableWidget.setRowCount(self.rows)
        # Встановлюємо заголовок на нашій таблиці.
        self.ui.tableWidget.setHorizontalHeaderLabels(["Студенти"])
        # Робимо розмір колонок, як наш віджет
        self.ui.tableWidget.setColumnWidth(1, 200)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        # Заповнюємо наші лейбли даними про групу.
        self.ui.label_5.setText(form[0])
        self.ui.label_6.setText(form[1])
        self.ui.label_7.setText(form[2])
        self.ui.label_8.setText(form[3])
        # Натискання кнопок Створити та Назад.
        self.ui.pushButton.clicked.connect(self.click_1)
        self.ui.pushButton_2.clicked.connect(self.click_2)
        # Показуэмо наше вікно.
        self.ui.show()

    def click_1(self):
        """
        Слот для видалення старих даних про групу, та створення нових файлів.
        """
        # Видаляємо файли зі старими данними.
        try:
            for i in range(15):
                os.remove(str(i + 1) + ' Тиждень.xls')
        except FileNotFoundError:
            pass
        # Створюємо новий файл з переліком студентів.
        try:
            for i in range(self.weeks):
                book = xlwt.Workbook(encoding="utf-8")
                sheet_1 = book.add_sheet('Week')
                sheet_1.write(0, 0, "Студенти/Пари")

                for j in range(self.rows):
                    text = self.ui.tableWidget.item(j, 0).text()
                    sheet_1.write(j + 1, 0, text)
                book.save(str(i + 1) + ' Тиждень.xls')

            DialogStudent(self)
        except AttributeError:
            self.msgWarn = QtWidgets.QMessageBox()
            self.msgWarn.setIcon(QtWidgets.QMessageBox.Warning)
            self.msgWarn.setText("Не всі комірки студентів заповнені.")
            self.msgWarn.setStandardButtons(QtWidgets.QMessageBox.Ok)
            self.msgWarn.buttonClicked.connect(self.msgWarn.close)
            self.msgWarn.exec_()

    def click_2(self):
        GroupData(self)
        self.ui.close()


class DialogStudent(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(DialogStudent, self).__init__(parent)
        self.ui = uic.loadUi('Діалог завершення створення журналу.ui')
        self.ui.setWindowTitle('Dialog')
        self.ui.pushButton.clicked.connect(parent.ui.close)
        self.ui.pushButton.clicked.connect(self.open_main_menu)
        self.ui.show()

    def open_main_menu(self):
        self.ui.close()
        MainWindow(self)


class WeekChoose(QtWidgets.QMainWindow):
    """
    Вікно для вибору та перегляду тижня, також тут можна побудувати графік пропусків, та подивитись рекомендації на
    відрахування.
    """
    def __init__(self, parent=None):
        super(WeekChoose, self).__init__(parent)
        self.ui = uic.loadUi('Вибір тижня.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        # Натискання на кнопку, щоб повернутись назад.
        self.ui.pushButton.clicked.connect(self.plot)
        self.ui.pushButton_2.clicked.connect(self.testimonial)
        self.ui.pushButton_3.clicked.connect(self.click_1)
        # Створюємо таблицю для вибору тижня.
        self.ui.tableWidget.setColumnCount(1)
        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        # Задаємо кількість тижнів в таблиці.
        self.ui.tableWidget.setRowCount(int(form[3]))
        # Заголовок Тижнів.
        self.ui.tableWidget.setHorizontalHeaderLabels(["Тижні"])
        # Робимо розмір наших клітинок, під розмір нашого табличного віджету.
        self.ui.tableWidget.setColumnWidth(1, 200)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        # Заповнюємо поля даними про групу.
        self.ui.label_5.setText(form[0])
        self.ui.label_6.setText(form[1])
        self.ui.label_7.setText(form[2])
        self.ui.label_8.setText(form[3])
        # Заповнюємо нашу таблицю назвами Тижнів.
        for i in range(int(form[3]) + 1):
            self.ui.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(str(i + 1) + " Тиждень"))
        self.ui.tableWidget.cellClicked.connect(self.click_on_week)
        self.ui.show()

    def click_1(self):
        """
        Слот для повернення назад.
        """
        MainWindow(self)
        self.ui.close()

    def plot(self):
        """
        Слот, для побудови графіків відвідування.
        """
        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        array = []
        df = pd.read_excel('1 Тиждень' + '.xls')
        data_names = []
        for i in df['Студенти/Пари']:
            data_names.append(i)
        for i in range(int(form[3])):
            data_list = []
            df = pd.read_excel(str(i + 1) + ' Тиждень' + '.xls')
            df = pd.DataFrame(df)
            try:
                for j in df['Unnamed: 49']:
                    data_list.append(j)
                array.append(data_list)
            except KeyError:
                continue
        array = np.array(array)
        data_frame = pd.DataFrame(array)
        data_values = []
        for i in range(int(form[2])):
            data_values.append(data_frame[i].sum())
        _, ax = plt.subplots(figsize=(30, 30))
        ax.bar(data_names, data_values, color='#539caf', align='center')
        ax.set_ylabel('')
        ax.set_xlabel('')
        ax.set_title('Графік відвідування студентів.')
        matplotlib.use('Qt5Agg')
        plt.show()

    def testimonial(self):
        Deduction(self)
        self.ui.close()

    def click_on_week(self, row, column):
        """
        Слот для відкриття вибраного тижня, та передачі туди ID.
        """
        item = self.ui.tableWidget.item(row, column)
        self.ID = item.text()
        # Передаємо ID в наступне вікно, щоб знати, яке вікно потрібно відкрити.
        Week(self).init_ui(self.ID)
        self.ui.close()


class Deduction(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Deduction, self).__init__(parent)
        self.ui = uic.loadUi('Рекомендації на відрахування.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        self.ui.pushButton_3.clicked.connect(self.click_1)
        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        array = []
        df = pd.read_excel('1 Тиждень' + '.xls')
        self.data_names = []
        for i in df['Студенти/Пари']:
            self.data_names.append(i)
        for i in range(int(form[3])):
            data_list = []
            df = pd.read_excel(str(i + 1) + ' Тиждень' + '.xls')
            df = pd.DataFrame(df)
            try:
                for j in df['Unnamed: 49']:
                    data_list.append(j)
                array.append(data_list)
            except KeyError:
                continue
        array = np.array(array)
        data_frame = pd.DataFrame(array)
        self.data_values = []
        for i in range(int(form[2])):
            self.data_values.append(data_frame[i].sum())

        d = {}
        for i in range(len(self.data_values)):
            d[self.data_names[i]] = self.data_values[i]
        self.deduction = []
        for i in d:
            if d[i] > int(form[3]) * 3:
                self.deduction.append(i)
        self.ui.tableWidget.setColumnCount(1)
        self.ui.tableWidget.setRowCount(len(self.deduction))
        self.ui.tableWidget.setHorizontalHeaderLabels(["Студенти"])
        # Робимо розмір колонок, як наш віджет
        self.ui.tableWidget.setColumnWidth(1, 200)
        self.ui.tableWidget.horizontalHeader().setSectionResizeMode(0, QtWidgets.QHeaderView.Stretch)
        for i in range(len(self.deduction)):
            self.ui.tableWidget.setItem(i, 0, QtWidgets.QTableWidgetItem(self.deduction[i]))
        self.ui.show()

    def click_1(self):
        WeekChoose(self)
        self.ui.close()


class Week(QtWidgets.QMainWindow):
    """
    Вікно таблиці тижня.
    """
    def __init__(self, parent=None):
        super(Week, self).__init__(parent)

    def init_ui(self, ID):
        self.ID = ID
        self.ui = uic.loadUi('Тижні.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        # Задаємо текст нашого Label.
        self.ui.label.setText(self.ID)
        # Натискання кнопок, щоб зберегти, вийти, або обрахувати загальну кількість пропусків.
        self.ui.pushButton_2.clicked.connect(self.click_1)
        self.ui.pushButton_3.clicked.connect(self.click_2)
        self.ui.pushButton.clicked.connect(self.click_3)
        self.ui.toolButton.setIcon(QtGui.QIcon('question-mark.png'))
        self.ui.toolButton_2.setIcon(QtGui.QIcon('calendar.png'))
        self.ui.toolButton_2.clicked.connect(self.calendar)
        self.ui.toolButton.clicked.connect(self.tips)
        # відкриваємо файл з даними про групу.
        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        # встановлюємо розмір нашої таблиці
        self.ui.tableWidget.setRowCount(int(form[2]) + 1)
        self.ui.tableWidget.setColumnCount(50)
        self.ui.tableWidget.verticalHeader().setVisible(False)
        # робимо заголовки в таблиці
        headers = ['Студенти']
        weekdays = ["Понеділок", "Вівторок", "Середа", "Четвер", "П'ятниця", "Субота"]
        for i in weekdays:
            for j in range(8):
                headers.append(i + ' ' + str(j + 1) + '. пара')
        headers.append('Всього пропусків')
        self.ui.tableWidget.setHorizontalHeaderLabels(headers)
        # встановлюємо в колонці заголовок Студенти/Пари
        newItem = QtWidgets.QTableWidgetItem("Студенти/Пари")
        self.ui.tableWidget.setItem(0, 0, newItem)
        # Заповнюємо таблицю даними з файлу excel.
        book = xlrd.open_workbook(self.ID + '.xls')
        sheet = book.sheet_by_name("Week")
        data = [[sheet.cell_value(r, c) for c in range(sheet.ncols)] for r in range(sheet.nrows)]
        for row, columnvalues in enumerate(data):
            for column, value in enumerate(columnvalues):
                item = QtWidgets.QTableWidgetItem(value)
                self.ui.tableWidget.setItem(row, column, item)

        self.ui.show()

    def click_1(self):
        """
        Слот, для повернення назад.
        """
        WeekChoose(self)
        self.ui.close()

    def click_2(self):
        """
        Слот, для збереження змін, внесених в таблицю.
        """
        try:
            num_rows = self.ui.tableWidget.rowCount()
            num_cols = self.ui.tableWidget.columnCount()

            tmp_df = pd.DataFrame(columns=range(num_cols), index=range(num_rows))
            for i in range(num_rows):
                for j in range(num_cols):
                    try:
                        tmp_df.iloc[i, j] = self.ui.tableWidget.item(i, j).text()
                    except AttributeError:
                        tmp_df.iloc[i, j] = 'None'

            del tmp_df[0]
            del tmp_df[49]
            for i in range(1, num_rows):
                for j in tmp_df.loc[i]:
                    assert j == '1' or j == '0' or j == 'None'

            with open('group.txt', 'r') as f:
                form = f.read().split('\n')
            book = xlwt.Workbook(encoding="utf-8")
            sheet_1 = book.add_sheet("Week")
            for i in range(int(form[2]) + 2):
                for j in range(501):
                    try:
                        text = self.ui.tableWidget.item(i, j).text()
                        sheet_1.write(i, j, text)
                    except AttributeError:
                        continue
            book.save(self.ID + '.xls')
            DialogSave(self)
        except AssertionError:
            DialogWarning(self)

    def click_3(self):
        """
        Слот, для обрахування загальної кількості пропусків студентів.
        """
        df = pd.read_excel(self.ID + '.xls')
        df = pd.DataFrame(df)
        del df['Студенти/Пари']
        try:
            del df['Unnamed: 49']
            df['All'] = df.sum(axis=1)
            for i in range(len(df['All'])):
                self.ui.tableWidget.setItem(i + 1, 49, QtWidgets.QTableWidgetItem(str(df['All'][i])))
        except KeyError:
            df['All'] = df.sum(axis=1)
            for i in range(len(df['All'])):
                self.ui.tableWidget.setItem(i + 1, 49, QtWidgets.QTableWidgetItem(str(df['All'][i])))

        with open('group.txt', 'r') as f:
            form = f.read().split('\n')
        book = xlwt.Workbook(encoding="utf-8")
        sheet_1 = book.add_sheet("Week")
        for i in range(int(form[2]) + 2):
            for j in range(501):
                try:
                    text = self.ui.tableWidget.item(i, j).text()
                    sheet_1.write(i, j, text)
                except AttributeError:
                    continue
        book.save(self.ID + '.xls')
        DialogCount(self)

    def tips(self):
        Tips(self)

    def calendar(self):
        Calendar(self)


class Tips(QtWidgets.QMainWindow):
    def __init__(self, parent=None):
        super(Tips, self).__init__(parent)
        self.ui = uic.loadUi('Підказки заповнення журналу.ui')
        self.ui.setWindowTitle('Електронний журнал старости')
        self.ui.pushButton_3.clicked.connect(self.ui.close)
        self.ui.show()


class DialogSave(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(DialogSave, self).__init__(parent)
        self.ui = uic.loadUi('Збереження змін.ui')
        self.ui.setWindowTitle('Dialog')
        self.ui.pushButton.clicked.connect(self.ui.close)
        self.ui.show()


class DialogCount(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(DialogCount, self).__init__(parent)
        self.ui = uic.loadUi('Обрахування пропусків.ui')
        self.ui.setWindowTitle('Dialog')
        self.ui.pushButton.clicked.connect(self.ui.close)
        self.ui.show()


class DialogWarning(QtWidgets.QDialog):
    def __init__(self, parent=None):
        super(DialogWarning, self).__init__(parent)
        self.ui = uic.loadUi('Не коректний запис.ui')
        self.ui.setWindowTitle('Dialog')
        self.ui.pushButton.clicked.connect(self.ui.close)
        self.ui.show()


class Calendar(QtWidgets.QWidget):
    def __init__(self, parent=None):
        super(Calendar, self).__init__(parent)
        self.ui = uic.loadUi('Календар.ui')
        self.ui.setWindowTitle('Електронний журнал старости.')
        self.ui.pushButton.clicked.connect(self.ui.close)
        self.ui.show()


if __name__ == '__main__':
    app = QtWidgets.QApplication(sys.argv)
    w = MainWindow()
    app.setWindowIcon(QtGui.QIcon('Hopstarter-Book-Book-Blank-Book.ico'))
    sys.exit(app.exec())
