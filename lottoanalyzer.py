import openpyxl

class LottoAnalyzer():
    def setData(self, filename):
        book = openpyxl.load_workbook(filename)
        sheet = book.active
        data_ragne = sheet['B4':'T804']

        lottoInfo = {}
        for row in data_ragne:
            values = []
            for cell in row:
                values.append(cell.value)

            try:
                lottodata = {}
                lottodata['추첨일'] = values[1]
                for i in range(5):
                    winnerdata = {}
                    winnerdata['당첨자수'] = values[i*2 + 2]
                    winnerdata['당첨금액'] = values[i*2 + 3]
                    lottodata[i+1] = winnerdata
                lottodata['당첨번호'] = values[12:19]
                lottoInfo[values[0]] = lottodata
            except:
                lottodata = {}
                lottodata['추첨일'] = values[1]
                for i in range(5):
                    winnerdata = {}
                    winnerdata['당첨자수'] = values[i*2 + 2]
                    winnerdata['당첨금액'] = values[i*2 + 3]
                    lottodata[i+1] = winnerdata
                lottodata['당첨번호'] = values[12:19]
                lottoInfo[values[0]] = lottodata

        return lottoInfo


if __name__ == "__main__":
    lottoAnalyzer = LottoAnalyzer()
    result = lottoAnalyzer.setData("lotto.xlsx")

    for index, info in result.items():
        print(index, info)

