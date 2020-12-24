import sys
import excel

from PyQt5 import uic
from PyQt5.QtWidgets import QApplication, QMainWindow, QFileDialog, QTableWidgetItem, QMessageBox

form_class = uic.loadUiType("ui/StandardIntegration.ui")[0]


class MyWindow(QMainWindow, form_class):
    def __init__(self):
        super().__init__()
        self.pageNum = 0
        self.setupUi(self)
        self.btnTrgStdFileOpen.clicked.connect(self.btnTrgStdFileOpenClicked)
        self.btnCmmnStdFileOpen.clicked.connect(self.btnCmmnStdFileOpenClicked)
        self.btnCompare.clicked.connect(self.compare_excel)
        self.btnIntegration.clicked.connect(self.integrate_excel)
        self.btnMakeFile.clicked.connect(self.makeCompareFile)
        self.btnMakeIntegratedFile.clicked.connect(self.makeIntegratedFile)

        self.btnGoPage1.clicked.connect(self.prevPage)
        self.btnGoPage2.clicked.connect(self.prevPage)

        self.firstPageExcelUploadInfo = {
            'trg': False,
            'cmmn': False
        }

        self.readyDownLoad = {
            'integration': False,
            'compare': False
        }

        self.stackedWidget.setCurrentIndex(self.pageNum)

    def nextPage(self):
        self.pageNum = self.pageNum + 1
        self.stackedWidget.setCurrentIndex(self.pageNum)

    def prevPage(self):
        self.pageNum = self.pageNum - 1
        self.stackedWidget.setCurrentIndex(self.pageNum)

    def integrate_excel(self):
        try:
            read = excel.Read()
            self.merge_term = read.merge_excel(read.std_excel(self.trgFile[0]),
                                               read.cmmn_std_excel(self.cmmnFile[0]))

            dataLen = len(self.merge_term)
            if dataLen > 0:
                self.integratedTable.setRowCount(dataLen)
                for idx, item in enumerate(self.merge_term):
                    self.integratedTable.setItem(idx, 0, QTableWidgetItem(self.getData(item, 'dv')))
                    self.integratedTable.setItem(idx, 1, QTableWidgetItem(self.getData(item, 'termNm')))
                    self.integratedTable.setItem(idx, 2, QTableWidgetItem(self.getData(item, 'termEngNm')))
                    self.integratedTable.setItem(idx, 3, QTableWidgetItem(self.getData(item, 'domNm')))
                    self.integratedTable.setItem(idx, 4, QTableWidgetItem(self.getData(item, 'datTp')))
                    self.integratedTable.setItem(idx, 5, QTableWidgetItem(self.getData(item, 'datLen')))
                    self.integratedTable.setItem(idx, 6, QTableWidgetItem(self.getData(item, 'datDcmlLen')))
                    self.integratedTable.setItem(idx, 7, QTableWidgetItem(self.getData(item, 'stdYn')))
                    self.integratedTable.setItem(idx, 8, QTableWidgetItem(self.getData(item, 'mappingTermNm')))
                    self.integratedTable.setItem(idx, 9, QTableWidgetItem(self.getData(item, 'mappingTermEngNm')))
                    self.integratedTable.setItem(idx, 10, QTableWidgetItem(self.getData(item, 'mappingDatTp')))
                    self.integratedTable.setItem(idx, 11, QTableWidgetItem(self.getData(item, 'mappingDatLen')))
                    self.integratedTable.setItem(idx, 12, QTableWidgetItem(self.getData(item, 'mappingDatDcmlLen')))
                    self.integratedTable.setItem(idx, 13, QTableWidgetItem(self.getData(item, 'desc')))
                self.readyDownLoad['integration'] = True
            else:
                self.readyDownLoad['integration'] = False
                # print('데이터가 없습니다.')
                self.err_msg_box('데이터가 없습니다.')

            self.nextPage()
        except Exception as e:  # 예외가 발생했을 때 실행됨
            self.readyDownLoad['integration'] = False
            # print('err', e)
            self.err_msg_box('엑셀파일 작성중에 오류가 발생하였습니다.')

    def compare_excel(self):

        if not self.firstPageExcelUploadInfo['trg'] or not self.firstPageExcelUploadInfo['cmmn']:
            err_msg = "기준데이터 표준과 공통표준파일이 모두 선택되어야 합니다."
            # print(err_msg)
            self.err_msg_box(err_msg)
        else:

            try:
                read = excel.Read()
                self.compare_term = read.compare_excel(read.std_excel(self.trgFile[0]),
                                                       read.cmmn_std_excel(self.cmmnFile[0]))

                dataLen = len(self.compare_term)
                if dataLen > 0:
                    self.compareTable.setRowCount(dataLen)
                    for idx, item in enumerate(self.compare_term):
                        self.compareTable.setItem(idx, 0, QTableWidgetItem(self.getData(item, 'termNm')))
                        self.compareTable.setItem(idx, 1, QTableWidgetItem(self.getData(item, 'termEngNm')))
                        self.compareTable.setItem(idx, 2, QTableWidgetItem(self.getData(item, 'datTp')))
                        self.compareTable.setItem(idx, 3, QTableWidgetItem(self.getData(item, 'datLen')))
                        self.compareTable.setItem(idx, 4, QTableWidgetItem(self.getData(item, 'datDcmlLen')))
                        self.compareTable.setItem(idx, 5, QTableWidgetItem(self.getData(item, 'termDesc')))
                        self.compareTable.setItem(idx, 6, QTableWidgetItem(self.getData(item, 'cmmnTermNm')))
                        self.compareTable.setItem(idx, 7, QTableWidgetItem(self.getData(item, 'cmmnTermEngNm')))
                        self.compareTable.setItem(idx, 8, QTableWidgetItem(self.getData(item, 'cmmnDatTp')))
                        self.compareTable.setItem(idx, 9, QTableWidgetItem(self.getData(item, 'cmmnDatLen')))
                        self.compareTable.setItem(idx, 10, QTableWidgetItem(self.getData(item, 'cmmnDatDcmlLen')))
                        self.compareTable.setItem(idx, 11, QTableWidgetItem(self.getData(item, 'stdYn')))
                    self.readyDownLoad['compare'] = True
                else:
                    self.readyDownLoad['compare'] = False
                    # print('데이터가 없습니다.')
                    self.err_msg_box('데이터가 없습니다.')

                self.nextPage()
            except Exception as e:  # 예외가 발생했을 때 실행됨
                self.readyDownLoad['compare'] = False
                # print('err', e)
                self.err_msg_box('엑셀파일 작성중에 오류가 발생하였습니다.')

    def makeCompareFile(self):

        if not self.readyDownLoad['compare']:
            err_msg = "비교 매핑된 데이터가 존재하지 않아 파일을 생성할 수 없습니다."
            # print(err_msg)
            self.err_msg_box(err_msg)

        else:
            fileNm = QFileDialog.getSaveFileName(self, 'Save file', "",
                                                 "Excel 통합 문서 (*.xlsx);; Excel 97-2003 통합 문서 (*.xlsx)")

            if fileNm != "":
                write = excel.Write()
                write.compare_excel(self.compare_term, fileNm[0])

    def makeIntegratedFile(self):
        if not self.readyDownLoad['integration']:
            err_msg = "통합된 데이터가 존재하지 않아 파일을 생성할 수 없습니다."
            # print(err_msg)
            self.err_msg_box(err_msg)

        else:
            fileNm = QFileDialog.getSaveFileName(self, 'Save file', "",
                                                 "Excel 통합 문서 (*.xlsx);; Excel 97-2003 통합 문서 (*.xlsx)")

            if fileNm != "":
                write = excel.Write()
                write.merge_excel(self.merge_term, fileNm[0])

    def btnTrgStdFileOpenClicked(self):
        self.trgFile = QFileDialog.getOpenFileName(self, 'Open file', "",
                                                   "All Files(*);; Excel 통합 문서 (*.xlsx);; Excel 97-2003 통합 문서 (*.xlsx)",
                                                   "")

        if self.trgFile[0] != "":
            self.txTrgStdFileNm.setText(self.trgFile[0])
            self.readTrgStdExcel()

    def btnCmmnStdFileOpenClicked(self):
        self.cmmnFile = QFileDialog.getOpenFileName(self, 'Open file', "",
                                                    "All Files(*);; Excel 통합 문서 (*.xlsx);; Excel 97-2003 통합 문서 (*.xlsx)",
                                                    "")
        if self.cmmnFile[0] != "":
            self.txCmmnStdFileNm.setText(self.cmmnFile[0])
            self.readCmmnStdExcel()

    def readCmmnStdExcel(self):
        try:
            read = excel.Read()

            self.cmmn_std = read.cmmn_std_excel(self.cmmnFile[0])
            # print('공통표준', self.cmmn_std)

            dataLen = len(self.cmmn_std)
            if dataLen > 0:
                self.cmmnStdTable.setRowCount(dataLen)
                for idx, item in enumerate(self.cmmn_std):
                    self.cmmnStdTable.setItem(idx, 0, QTableWidgetItem(self.getData(item, 'termNm')))
                    self.cmmnStdTable.setItem(idx, 1, QTableWidgetItem(self.getData(item, 'termEngNm')))
                    self.cmmnStdTable.setItem(idx, 2, QTableWidgetItem(self.getData(item, 'domNm')))
                    self.cmmnStdTable.setItem(idx, 3, QTableWidgetItem(self.getData(item, 'datTp')))
                    self.cmmnStdTable.setItem(idx, 4, QTableWidgetItem(self.getData(item, 'datLen')))
                    self.cmmnStdTable.setItem(idx, 5, QTableWidgetItem(self.getData(item, 'datDcmlLen')))
                    self.cmmnStdTable.setItem(idx, 6, QTableWidgetItem(self.getData(item, 'termDesc')))
                self.firstPageExcelUploadInfo['cmmn'] = True
            else:
                self.firstPageExcelUploadInfo['cmmn'] = False
                # print('데이터가 없습니다.')
                self.err_msg_box('데이터가 없습니다.')
        except Exception as e:  # 예외가 발생했을 때 실행됨
            self.firstPageExcelUploadInfo['cmmn'] = False
            # print('err', e)
            self.err_msg_box('엑셀양식을 확인하세요.')

    def readTrgStdExcel(self):
        try:
            read = excel.Read()
            self.trg_std = read.std_excel(self.trgFile[0])
            # print('기관표준', self.trg_std)

            dataLen = len(self.trg_std)
            if dataLen > 0:
                self.trgStdTable.setRowCount(dataLen)
                for idx, item in enumerate(self.trg_std):
                    self.trgStdTable.setItem(idx, 0, QTableWidgetItem(self.getData(item, 'termNm')))
                    self.trgStdTable.setItem(idx, 1, QTableWidgetItem(self.getData(item, 'termEngNm')))
                    self.trgStdTable.setItem(idx, 2, QTableWidgetItem(self.getData(item, 'datTp')))
                    self.trgStdTable.setItem(idx, 3, QTableWidgetItem(self.getData(item, 'datLen')))
                    self.trgStdTable.setItem(idx, 4, QTableWidgetItem(self.getData(item, 'datDcmlLen')))
                    self.trgStdTable.setItem(idx, 5, QTableWidgetItem(self.getData(item, 'termDesc')))
                self.firstPageExcelUploadInfo['trg'] = True
            else:
                self.firstPageExcelUploadInfo['trg'] = False
                # print('데이터가 없습니다.')
                self.err_msg_box('데이터가 없습니다.')

        except Exception as e:  # 예외가 발생했을 때 실행됨
            self.firstPageExcelUploadInfo['trg'] = False
            # print('err', e)
            self.err_msg_box('엑셀양식을 확인하세요.')

    def getData(self, item, keyStr):
        if keyStr in item:
            if not item[keyStr]:
                return ''
            else:
                return item[keyStr]
        else:
            return ''

    def err_msg_box(self, msg):
        QMessageBox.critical(self, "오류", msg)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = MyWindow()
    window.show()
    app.exec_()
