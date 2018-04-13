import openpyxl
import sqlite3

class LottoAnalyzer():
    lottoInfo = {}

    def setData(self, filename):
        book = openpyxl.load_workbook(filename)
        sheet = book.active
        data_ragne = sheet['B4':'T804']

        #for row in data_ragne:
        for row in range(len(data_ragne)-1, -1, -1):
            values = []
            for cell in data_ragne[row]:
                values.append(cell.value)

            try:
                lottodata = {}
                lottodata['추첨일'] = values[1]
                for i in range(5):
                    winnerdata = {}
                    winnerdata['당첨자'] = int(values[i*2 + 2])
                    winnerdata['당첨금'] = int(values[i*2 + 3][:-1].replace(',',''))
                    lottodata[i+1] = winnerdata
                lottodata['당첨번호'] = values[12:19]
                self.lottoInfo[values[0]] = lottodata
            except:
                lottodata = {}
                lottodata['추첨일'] = values[1]
                for i in range(5):
                    winnerdata = {}
                    winnerdata['당첨자'] = int(values[i*2 + 2])
                    winnerdata['당첨금'] = int(values[i*2 + 3][:-1].replace(',',''))
                    lottodata[i+1] = winnerdata
                lottodata['당첨번호'] = values[12:19]
                self.lottoInfo[values[0]] = lottodata

        return self.lottoInfo

    def getTotalNum(self):
        return len(self.lottoInfo)

    def getDateList(self):
        dateList = []
        for index, value in self.lottoInfo.items():
            dateList.append(value['추첨일'])
        return dateList

    def savetoDB(self, path, data):
        con = sqlite3.connect(path)
        cursor = con.cursor()
        cursor.execute("CREATE TABLE lotto(회차 int, 추첨일 text, 당첨자1 int, 당첨금1 int, 당첨자2 int, 당첨금2 int, \
            당첨자3 int, 당첨금3 int, 당첨자4 int, 당첨금4 int, 당첨자5 int, 당첨금5 int, 번호1 int, 번호2 int, \
            번호3 int, 번호4 int, 번호5 int, 번호6 int, 보너스 int)")

        for index, info in data.items():
            cursor.execute("INSERT INTO lotto VALUES(?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)", \
                           (index, info['추첨일'], info[1]['당첨자'], info[1]['당첨금'], \
                            info[2]['당첨자'], info[2]['당첨금'], \
                            info[3]['당첨자'], info[3]['당첨금'], \
                            info[4]['당첨자'], info[4]['당첨금'], \
                            info[5]['당첨자'], info[5]['당첨금'], \
                            info['당첨번호'][0], info['당첨번호'][1], info['당첨번호'][2], \
                            info['당첨번호'][3], info['당첨번호'][4], info['당첨번호'][5], \
                            info['당첨번호'][6]))
        con.commit()
        con.close()


if __name__ == "__main__":
    lottoAnalyzer = LottoAnalyzer()
    result = lottoAnalyzer.setData("lotto.xlsx")

    #for index, info in result.items():
        #print(index, info)
    lottoAnalyzer.savetoDB("./lotto.db", result)

    #print(lottoAnalyzer.getTotalNum())
    #print(lottoAnalyzer.getDateList())
