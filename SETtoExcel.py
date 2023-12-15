import requests
import datetime
import xlsxwriter
import os
import json

from seleniumwire import webdriver
from seleniumwire.utils import decode

def main():
    # list of api path (use to identify stock data)
    SET_path = ["/api/set/index/AGRO/composition", 
                "/api/set/index/CONSUMP/composition", 
                "/api/set/index/FINCIAL/composition", 
                "/api/set/index/INDUS/composition", 
                "/api/set/index/PROPCON/composition",
                "/api/set/index/RESOURC/composition",
                "/api/set/index/SERVICE/composition",
                "/api/set/index/TECH/composition"]
    
    mai_path = ["/api/set/index/AGRO-m/composition",
                "/api/set/index/CONSUMP-m/composition",
                "/api/set/index/FINCIAL-m/composition",
                "/api/set/index/INDUS-m/composition",
                "/api/set/index/PROPCON-m/composition",
                "/api/set/index/RESOURC-m/composition",
                "/api/set/index/SERVICE-m/composition",
                "/api/set/index/TECH-m/composition"]
    
    # get current date
    now = datetime.datetime.now()
    hour = twoDigitsNum(now.hour)
    minute = twoDigitsNum(now.minute)
    year = now.year
    monthNumber = twoDigitsNum(now.month)
    monthName = monthNumToName(now.month)
    day = twoDigitsNum(now.day)

    # create lists of content (use to feed data to each column)
    SET_symbol = []
    mai_symbol = []
    SET_sign = []
    mai_sign = []
    SET_prior = []
    mai_prior = []
    SET_last = []
    mai_last = []
    SET_percentChange = []
    mai_percentChange = []
    SET_open = []
    mai_open = []
    SET_high = []
    mai_high = []
    SET_low = []
    mai_low = []
    SET_average = []
    mai_average = []
    SET_change = []
    mai_change = []
    SET_high52Weeks = []
    mai_high52Weeks = []
    SET_low52Weeks = []
    mai_low52Weeks = []
    SET_totalVolume = []
    mai_totalVolume = []
    SET_aomVolume = []
    mai_aomVolume = []
    SET_totalValue = []
    mai_totalValue = []
    SET_aomValue = []
    mai_aomValue = []
    SET_OBT_BigLot = []
    mai_OBT_BigLot = []
    SET_marketCap = []
    mai_marketCap = []
    SET_pbRatio = []
    mai_pbRatio = []
    SET_dividend = []
    mai_dividend = []
    SET_nvdrVolume = []
    mai_nvdrVolume = []
    SET_industryName = []
    mai_industryName = []
    SET_sectorName = []
    mai_sectorName = []

    SET_header = []
    mai_header = []

    #join ca.crt and ca.key
    selenium_wire_storage = os.path.join(os.getcwd(), "selenium_wire")

    # setting up the webdriver
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("--headless")
    seleniumwireOptions = {"disable_encoding": "True", "request_storage_base_dir": selenium_wire_storage}
    driver = webdriver.Chrome(options=chromeOptions, seleniumwire_options=seleniumwireOptions)

    # request SET data
    driver.get("https://www.settrade.com/th/equities/market-data/overview?category=Index&index=SET")

    for request in driver.requests:
        if request.response:
            if request.path in SET_path:
                response = request.response
                body = decode(response.body, response.headers.get("content-encoding", "identity"))
                body_str = body.decode()
                dataDict = json.loads(body_str)
                subIndices_list = dataDict["composition"]["subIndices"]
                
                # insert data into lists
                for sector in subIndices_list:
                    if sector["stockInfos"]:
                        for stock in sector["stockInfos"]:
                            SET_symbol.append(stock["symbol"])
                            SET_sign.append(stock["sign"])
                            SET_prior.append(twoDecimal(stock["prior"]))
                            SET_last.append(twoDecimal(stock["last"]))
                            SET_percentChange.append(percentSign(twoDecimal(stock["percentChange"])))
                            SET_open.append(twoDecimal(stock["open"]))
                            SET_high.append(twoDecimal(stock["high"]))
                            SET_low.append(twoDecimal(stock["low"]))
                            SET_average.append(twoDecimal(stock["average"]))
                            SET_change.append(twoDecimal(stock["change"]))
                            SET_high52Weeks.append(twoDecimal(stock["high52Weeks"]))
                            SET_low52Weeks.append(twoDecimal(stock["low52Weeks"]))
                            SET_totalVolume.append(twoDecimal(unitK(stock["totalVolume"])))
                            SET_aomVolume.append(twoDecimal(unitK(stock["aomVolume"])))

                            totalValue = twoDecimal(unitK(stock["totalValue"]))
                            SET_totalValue.append(totalValue)

                            aomValue = twoDecimal(unitK(stock["aomValue"]))
                            SET_aomValue.append(aomValue)

                            if totalValue == "-" or aomValue == "-":
                                SET_OBT_BigLot.append("-")
                            else:
                                SET_OBT_BigLot.append(twoDecimal(unitK(totalValue - aomValue)))
                            SET_marketCap.append(twoDecimal(unitM(stock["marketCap"])))
                            SET_pbRatio.append(twoDecimal(stock["pbRatio"]))
                            SET_dividend.append(percentSign(twoDecimal(stock["dividendYield"])))
                            SET_nvdrVolume.append(stock["nvdrNetVolume"])
                            SET_industryName.append(stock["industryName"])
                            SET_sectorName.append(stock["sectorName"])
                        else:
                            continue

    driver.quit()

    

    # setting up the webdriver
    chromeOptions = webdriver.ChromeOptions()
    chromeOptions.add_argument("--headless")
    selenium_wire_storage = os.path.join(os.getcwd(), "selenium_wire")
    seleniumwireOptions = {"disable_encoding": "True", "request_storage_base_dir": selenium_wire_storage}
    driver = webdriver.Chrome(options=chromeOptions, seleniumwire_options=seleniumwireOptions)

    # request mai data
    driver.get("https://www.settrade.com/th/equities/market-data/overview?category=Index&market=mai&index=mai")

    for request in driver.requests:
        if request.response:
            if request.path in mai_path:
                response = request.response
                body = decode(response.body, response.headers.get("content-encoding", "identity"))
                body_str = body.decode()
                dataDict = json.loads(body_str)

                # insert data into lists
                for stock in dataDict["composition"]["stockInfos"]:
                    if stock is not None:
                        mai_symbol.append(stock["symbol"])
                        mai_sign.append(stock["sign"])
                        mai_prior.append(twoDecimal(stock["prior"]))
                        mai_last.append(twoDecimal(stock["last"]))
                        mai_percentChange.append(percentSign(twoDecimal(stock["percentChange"])))
                        mai_open.append(twoDecimal(stock["open"]))
                        mai_high.append(twoDecimal(stock["high"]))
                        mai_low.append(twoDecimal(stock["low"]))
                        mai_average.append(twoDecimal(stock["average"]))
                        mai_change.append(twoDecimal(stock["change"]))
                        mai_high52Weeks.append(twoDecimal(stock["high52Weeks"]))
                        mai_low52Weeks.append(twoDecimal(stock["low52Weeks"]))
                        mai_totalVolume.append(twoDecimal(unitK(stock["totalVolume"])))
                        mai_aomVolume.append(twoDecimal(unitK(stock["aomVolume"])))

                        totalValue = twoDecimal(unitK(stock["totalValue"]))
                        mai_totalValue.append(totalValue)

                        aomValue = twoDecimal(unitK(stock["aomValue"]))
                        mai_aomValue.append(aomValue)

                        if totalValue == "-" or aomValue == "-":
                            mai_OBT_BigLot.append("-")
                        else:
                            mai_OBT_BigLot.append(twoDecimal(unitK(totalValue - aomValue)))
                        mai_marketCap.append(twoDecimal(unitM(stock["marketCap"])))
                        mai_pbRatio.append(twoDecimal(stock["pbRatio"]))
                        mai_dividend.append(percentSign(twoDecimal(stock["dividendYield"])))
                        mai_nvdrVolume.append(stock["nvdrNetVolume"])
                        mai_industryName.append(stock["industryName"])
                        mai_sectorName.append(stock["sectorName"])

    # append list of headers into a single list (use this list to loop)
    SET_header.append(SET_symbol)
    SET_header.append(SET_sign)
    SET_header.append(SET_prior)
    SET_header.append(SET_last)
    SET_header.append(SET_percentChange)
    SET_header.append(SET_open)
    SET_header.append(SET_high)
    SET_header.append(SET_low)
    SET_header.append(SET_average)
    SET_header.append(SET_change)
    SET_header.append(SET_high52Weeks)
    SET_header.append(SET_low52Weeks)
    SET_header.append(SET_totalVolume)
    SET_header.append(SET_aomVolume)
    SET_header.append(SET_totalValue)
    SET_header.append(SET_aomValue)
    SET_header.append(SET_OBT_BigLot)
    SET_header.append(SET_marketCap)
    SET_header.append(SET_pbRatio)
    SET_header.append(SET_dividend)
    SET_header.append(SET_nvdrVolume)
    SET_header.append(SET_industryName)
    SET_header.append(SET_sectorName)
    
    mai_header.append(mai_symbol)
    mai_header.append(mai_sign)
    mai_header.append(mai_prior)
    mai_header.append(mai_last)
    mai_header.append(mai_percentChange)
    mai_header.append(mai_open)
    mai_header.append(mai_high)
    mai_header.append(mai_low)
    mai_header.append(mai_average)
    mai_header.append(mai_change)
    mai_header.append(mai_high52Weeks)
    mai_header.append(mai_low52Weeks)
    mai_header.append(mai_totalVolume)
    mai_header.append(mai_aomVolume)
    mai_header.append(mai_totalValue)
    mai_header.append(mai_aomValue)
    mai_header.append(mai_OBT_BigLot)
    mai_header.append(mai_marketCap)
    mai_header.append(mai_pbRatio)
    mai_header.append(mai_dividend)
    mai_header.append(mai_nvdrVolume)
    mai_header.append(mai_industryName)
    mai_header.append(mai_sectorName)

    # create folder for excel files (if not already existed)
    excelFolder = os.path.relpath(rf"excel/{year}/{monthNumber}-{monthName}/")
    if not os.path.exists(excelFolder):
        os.makedirs(excelFolder)

    # create excel file
    excelPath = rf"excel/{year}/{monthNumber}-{monthName}/{year}-{monthNumber}-{day}T{hour}{minute}.xlsx"
    workbook = xlsxwriter.Workbook(excelPath)
    SET_sheet = workbook.add_worksheet("SET")
    mai_sheet = workbook.add_worksheet("mai")

    # write header first
    header_row = ["Symbol", "Sign", "Prior", "Last", "%Change", "Open", "High", "Low", "Average", "Change",
                  "High52Weeks", "Low52Weeks", "TotalVolume (k)", "aomVolume (k)", "totalValue (k.THB)", "aomValue (k.THB)",
                  "Off-Board/BigLot Trading (k.THB)", "MarketCap", "pbRatio", "%Dividend", "nvdrNetVolume", "IndustryName", "SectorName"]
    
    col = 0
    for header_item in header_row:
        SET_sheet.write(0, col, header_item)
        mai_sheet.write(0, col, header_item)
        col += 1

    # write SET data
    row = 1
    column = 0
    for header in SET_header:
        for item in header:
            SET_sheet.write(row, column, item)
            row += 1
        row = 1
        column += 1

    # write mai data
    row = 1
    column = 0
    for header in mai_header:
        for item in header:
            mai_sheet.write(row, column, item)
            row += 1
        row = 1
        column += 1

    # close()
    workbook.close()
    driver.quit()

    del driver.requests

########## END OF main() ##########

def monthNumToName(monthInt):
    return { 1: "January",
    2: "February",
    3: "March",
    4: "April",
    5: "May",
    6: "June",
    7: "July",
    8: "August",
    9: "September",
    10: "October",
    11: "November",
    12: "December"}[monthInt]

# "day" and "month" number of value < 10 to two-digits
def twoDigitsNum(number):
    if number < 10:
        number = f"0{number}"
    return number

def twoDecimal(number):
    if number == None or number == "-":
        return "-"
    else:
        return round(float(number), 2)

def percentSign(number):
    if number == None or number == "":
        return "-"
    else:
        return str(number) + "%"

def unitK(number):
    if number == None:
        return "-"
    else:
        return number / 1000
    
def unitM(number):
    if number == None:
        return "-"
    else:
        return number / 1000000

###################################

if __name__ == "__main__":
    main()