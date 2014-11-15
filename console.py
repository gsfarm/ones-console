import os
import sys
import datetime

from PySide.QtCore import Slot, Qt
from PySide.QtGui import QMainWindow, QApplication, QComboBox, QPlainTextEdit, QDockWidget, \
    QSplitter, QTextEdit, QTableView, QStandardItemModel, QSizePolicy, QStandardItem

import pythoncom
import win32com.client


class ConsoleWidget(QMainWindow):
    def __init__(self):
        super(ConsoleWidget, self).__init__()
        self.setWindowTitle('1c query')

        self._connection = None

        self._home = os.path.expanduser('~/%s' % QApplication.applicationName())
        if not os.path.isdir(self._home):
            os.mkdir(self._home)

        self.queryToolBar = self.addToolBar('Query')
        self.queryAction = self.queryToolBar.addAction('Run', self.executeQuery)
        self.queryAction.setDisabled(True)

        uri_history = list()
        path = os.path.join(self._home, 'uri_history.txt')
        if os.path.isfile(path):
            uri_history = open(path, 'r').read().split('\n')

        self.connectionToolBar = self.addToolBar('Connection')
        self.connectionUriCombo = QComboBox(self)
        self.connectionUriCombo.setEditable(True)
        if not uri_history:
            self.connectionUriCombo.addItem('File="";usr="";pwd="";')
            self.connectionUriCombo.addItem('Srvr="{host}";Ref="{ref}";Usr="{user}";Pwd="{password}";')
        else:
            self.connectionUriCombo.addItems(uri_history)
            self.connectionUriCombo.setCurrentIndex(len(uri_history) - 1)
        self.connectionUriCombo.setSizePolicy(QSizePolicy.MinimumExpanding, QSizePolicy.Maximum)
        self.connectionToolBar.addWidget(self.connectionUriCombo)

        self.onesVersionCombo = QComboBox(self)
        self.onesVersionCombo.addItems(['8.3', '8.2', '8.1', '8.0'])
        self.onesVersionCombo.setCurrentIndex(0)
        self.connectionToolBar.addWidget(self.onesVersionCombo)
        self.connectAction = self.connectionToolBar.addAction('Connect', self.connectOneS)
        self.disconnectAction = self.connectionToolBar.addAction('Disconnect', self.disconnectOneS)
        self.disconnectAction.setDisabled(True)

        self.logEdit = QPlainTextEdit(self)
        self.logDock = QDockWidget('Log', self)
        self.logDock.setWidget(self.logEdit)
        self.addDockWidget(Qt.BottomDockWidgetArea, self.logDock, Qt.Horizontal)

        self.splitter = QSplitter(Qt.Vertical, self)
        self.setCentralWidget(self.splitter)

        self.sqlEdit = QTextEdit(self)
        self.sqlEdit.setLineWrapMode(QTextEdit.NoWrap)

        path = os.path.join(self._home, 'last-sql.txt')
        if os.path.isfile(path):
            sql = open(path, 'r').read()
            self.sqlEdit.setText(sql)

        self.model = QStandardItemModel(self)
        self.tableView = QTableView(self)
        self.tableView.setModel(self.model)

        self.splitter.addWidget(self.sqlEdit)
        self.splitter.addWidget(self.tableView)
        self.splitter.setStretchFactor(0, 3)
        self.splitter.setStretchFactor(1, 2)

    def query(self, sql):
        if not self._connection:
            self.logEdit.appendPlainText('No connection')
            return None

        try:
            query = self._connection.NewObject('Query', sql)
            result = query.Execute()
        except Exception as e:
            self.logEdit.appendPlainText(str(e))
            return None

        return result

    def refresh(self, result):
        self.model.clear()

        columns = list()
        result_columns = result.Columns
        for index in range(result_columns.Count()):
            name = result_columns.Get(index).Name
            columns.append(name)

        self.model.setColumnCount(len(columns))
        for section, name in enumerate(columns):
            self.model.setHeaderData(section, Qt.Horizontal, name)

        select = result.Choose()
        self.logEdit.appendPlainText('Selected %d records' % select.Count())
        while select.Next():
            items = list()
            for index in range(len(columns)):
                value = select.Get(index)

                item = QStandardItem('')
                if isinstance(value, bool):
                    item.setText(value and 'Yes' or 'No')

                elif isinstance(value, (int, str)):
                    item.setText(str(value))

                elif isinstance(value, datetime.datetime):
                    item.setText(value.strftime('%Y.%m.%d %H:%M:%S'))

                else:
                    item.setText(str(value))
                items.append(item)

            self.model.appendRow(items)

    @Slot()
    def executeQuery(self):
        sql = self.sqlEdit.toPlainText()
        result = self.query(sql)
        if result:
            path = os.path.join(self._home, 'last-sql.txt')
            open(path, 'w').write(sql)
            self.refresh(result)

    @Slot()
    def connectOneS(self):
        uri = self.connectionUriCombo.currentText().strip()
        if not uri:
            self.logEdit.appendPlainText('Need a connection string')
            return

        version = self.onesVersionCombo.currentText()
        comName = "V%s.COMConnector" % str(version).replace('.', '')

        pythoncom.CoInitialize()
        try:
            obj = win32com.client.Dispatch(comName)
            self._connection = obj.Connect(uri)
        except Exception as e:
            self.logEdit.appendPlainText(str(e))
            return

        self.connectAction.setDisabled(True)
        self.disconnectAction.setEnabled(True)
        self.queryAction.setEnabled(True)

        uri_history = list()
        for i in range(self.connectionUriCombo.count()):
            uri_history.append(self.connectionUriCombo.itemText(i))

        if uri not in uri_history:
            self.connectionUriCombo.clearEditText()
            self.connectionUriCombo.addItem(uri)
            self.connectionUriCombo.setCurrentIndex(len(uri_history))
            uri_history.append(uri)
            path = os.path.join(self._home, 'uri_history.txt')
            open(path, 'w').write('\n'.join(uri_history))

    @Slot()
    def disconnectOneS(self):
        pythoncom.CoUninitialize()
        self._connection = None
        self.connectAction.setEnabled(True)
        self.disconnectAction.setDisabled(True)
        self.queryAction.setDisabled(True)

if __name__ == '__main__':
    QApplication.setApplicationName('1CConsole')
    application = QApplication(sys.argv)

    mainWidget = ConsoleWidget()
    mainWidget.show()

    sys.exit(application.exec_())