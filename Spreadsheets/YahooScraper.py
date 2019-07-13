import requests
from bs4 import BeautifulSoup

debug = True

def grabYahooFinanceDataRows():
    yahooPage = requests.get("https://finance.yahoo.com/quote/AMD/history?p=AMD")
    print("status code: {}".format(yahooPage.status_code))

    yahooSoup = BeautifulSoup(yahooPage.content, 'html.parser')
    rowlist = yahooSoup.findAll("tr",{"class":"BdT Bdc($c-fuji-grey-c) Ta(end) Fz(s) Whs(nw)"})
    print("Number of days: {}".format(len(rowlist)))
    yahooFinanceDataRows = []
    for day in rowlist:
        tablelist = day.findAll("td")
        dayDate = tablelist[0].text
        dayOpen = tablelist[1].text
        dayHigh = tablelist[2].text
        dayLow = tablelist[3].text
        dayClose = tablelist[4].text
        dayAdjClose = tablelist[5].text
        dayVolume = tablelist[6].text
        dataRow = []
        dataRow.append(dayDate)
        dataRow.append(dayOpen)
        dataRow.append(dayHigh)
        dataRow.append(dayLow)
        dataRow.append(dayClose)
        dataRow.append(dayAdjClose)
        dataRow.append(dayVolume)
        yahooFinanceDataRows.append(dataRow)

        if debug == True:
            print("")
            print("dayDate: {}".format(dayDate))
            print("Num of td elements in this day: {}".format(len(tablelist)))
            print(" dayOpen: {}".format(dayOpen))
            print(" dayHigh: {}".format(dayHigh))
            print(" dayLow: {}".format(dayLow))
            print(" dayClose: {}".format(dayClose))
            print(" dayAdjClose: {}".format(dayAdjClose))
            print(" dayVolume: {}".format(dayVolume))
            print(" dataRow length: {}".format(len(dataRow)))
            print(" yahooFinanceDataRows length: {}".format(len(yahooFinanceDataRows)))
            print()
    
    print("Data array created from {} to {}!".format(yahooFinanceDataRows[0][0],yahooFinanceDataRows[-1][0]))
    return yahooFinanceDataRows

if __name__ == "__main__":
    grabYahooFinanceDataRows()