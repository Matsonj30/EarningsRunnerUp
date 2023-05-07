import requests
import pandas as pd
from openpyxl import load_workbook
from datetime import date


#startLine()
#returns the right most column in the excel sheet which is blank
#this is where we will start inputting new tickers found for the day
#PARAMETERS: currentPage -> the excel sheet we want to find the starting point of
#RETURNS: NONE
def startLine(currentPage):
    rowNumber = 1
    while currentPage.cell(row = rowNumber, column = 1).value != None:
        rowNumber += 1
    return rowNumber


#writeToExcel()
#using data retrieved by finvizData() will write to excel sheet
#PARAMETERS: data -> data scraped from Finviz, excelPage -> which page we want to write to
#RETURNS: NONE
def writeToExcel(data, excelPage):
    excelFile = load_workbook("D:/Programming/Repositories/finvizTracker/data.xlsx")
    currentPage = excelFile[excelPage]

    startingLine = startLine(currentPage)
    index = 1
    for ticker in data[1]: #for each row we have in our data aka each ticker
        currentPage.cell(row = startingLine, column = 1).value = data[0][index] ##iterate each thing here
        currentPage.cell(row = startingLine, column = 2).value = data[1][index] ##iterate each thing here
        currentPage.cell(row = startingLine, column = 3).value = date.today() ##iterate each thing here
        currentPage.cell(row = startingLine, column = 4).value = data[3][index] ##iterate each thing here
        currentPage.cell(row = startingLine, column = 5).value = data[4][index] ##iterate each thing here
        currentPage.cell(row = startingLine, column = 6).value = data[5][index] ##iterate each thing here
        currentPage.cell(row = startingLine, column = 7).value = data[6][index] ##iterate each thing here
        index += 1
        startingLine += 1
    excelFile.save(("D:/Programming/Repositories/screenerSettings/highVolume.xlsx"))

#parseData()
#retrieves finviz screener data by utilizing pandas to read html tables on finviz.com
#after gaining the data we want, we put it in an array to pass to writeToExcel()
#PARAMETERS: url -> Finviz page we want to scrape, excelPage -> which page we want to write to
#RETURNS: NONE
def parseData(url, excelPage):
    header = {'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_10_1) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/39.0.2171.95 Safari/537.36'}
    finvizPage = requests.get(url, headers=header).text  
    tables = pd.read_html(finvizPage)
    table = tables[-2] #this is the table we want from the many tables pandas found
    names = table.iloc[1:, 1] #I think _: is better than 1: even if do the same thing
    industries = table.iloc[1:, 4]
    marketCaps = table.iloc[1:,6]
    prices = table.iloc[1:,8]
    changes = table.iloc[1:,9]
    volumes = table.iloc[1,10] #[row selection, column selection] so 1:,10 means every row, but only column 10

    data = [names, industries, marketCaps, prices, changes, volumes]
    writeToExcel(data, excelPage)

parseData('https://finviz.com/screener.ashx?v=111&f=sh_float_u100,sh_relvol_o10&ft=4&o=volume', 'IrregularVolume')