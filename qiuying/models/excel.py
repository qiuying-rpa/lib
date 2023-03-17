class ExcelObj:
    def __init__(self, xlapp, workbook, worksheet):
        self._xlApp = xlapp
        self._workbook = workbook
        self._worksheet = worksheet

    @property
    def xlapp(self):
        return self._xlApp

    @property
    def workbook(self):
        return self._workbook

    @property
    def worksheet(self):
        return self._worksheet

    @worksheet.setter
    def worksheet(self, value):
        self._worksheet = value
