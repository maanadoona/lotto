import openpyxl
import sqlite3

class LottoAnalyzer():
    lottoInfo = {}

    def readFromExcel(self, filename):
        book = openpyxl.load_workbook(filename)
        sheet = book.active
        startPos = 'B4'
        endPos = 'T' + str(sheet.max_row)
        data_range = sheet[startPos:endPos]

        for row in range(len(data_range)-1, -1, -1):
            values = []
            for cell in data_range[row]:
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

    def saveToDB(self, path, data):
        con = sqlite3.connect(path)
        cursor = con.cursor()
        cursor.execute("CREATE TABLE if not exists lotto(회차 int, 추첨일 text, 당첨자1 int, 당첨금1 int, 당첨자2 int, 당첨금2 int, \
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

    def readFromDB(self, db, table):
        con = sqlite3.connect(db)
        cursor = con.cursor()

        query = "select * from {0}".format(table)
        cursor.execute(query)
        rows = cursor.fetchall()

        con.close()

        return rows

if __name__ == "__main__":
    lottoAnalyzer = LottoAnalyzer()

    # 기존 DB 로드
    dbData = lottoAnalyzer.readFromDB("./lotto.db", "lotto")
    #for d in dbData:
        #print(d)

    # 입력 데이터
    inputData = lottoAnalyzer.readFromExcel("lotto.xlsx")
    #for key, data in inputData.items():
       #print(key, data)

    # 기존과 입력 비교하여 저장할 데이터 추출
    toSaveNum = len(inputData) - len(dbData)
    print(toSaveNum)

    toSaveData = {}
    for i in range(len(inputData) - toSaveNum+1, len(inputData)+1):
        toSaveData[i] = inputData[i]

    for j in toSaveData:
        print(toSaveData[j])

    lottoAnalyzer.saveToDB("./lotto.db", toSaveData)

    #for index, info in result.items():
        #print(index, info)
    #lottoAnalyzer.saveToDB("./lotto.db", result)


    #print(lottoAnalyzer.getTotalNum())
    #print(lottoAnalyzer.getDateList())
