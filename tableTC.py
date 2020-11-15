import xlsxwriter
import testData

workbook = xlsxwriter.Workbook('TableFunction.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, 'Test Case')
worksheet.write(0, 1, 'Steps')
worksheet.write(0, 2, 'Expected Result')

def enterUrl(name, row, col):
    worksheet.write(row, col + 1, 'Login to ' + name)
    worksheet.write(row, col + 2, 'Go to ' + testData.startUrl)

def sideBar(name, row, col):
    worksheet.write(row, col + 1, 'Click ' + name + ' sidebar')
    worksheet.write(row, col + 2, name + ' page should be displayed')

def column(name, row, col):
    worksheet.write(row, col + 1, 'Check ' + name + ' column')
    worksheet.write(row, col + 2, name + ' page should be displayed')

def sortColumn(name, row, col):
    worksheet.write(row, col + 1, 'Click ' + name + ' column')
    worksheet.write(row, col + 2, name + ' column should be sorted')

row = 1
col = 0

#Check column
for name, identification in (testData.testcase):
    if (identification == 'url'):
        urlName = name
        urlTrue = True
    if (identification == 'sideBar'):
        sideBarName = name
        sideBarTrue = True
    if (identification == 'column'):
        worksheet.write(row, 0, 'Check ' + name + ' column')
        if (urlTrue):
            enterUrl(urlName, row, col)
            row += 1
        if (sideBarTrue):
            sideBar(sideBarName, row, col)
            row += 1
        column(name, row, col)
        row += 1

#Check column sorting
for name, identification in (testData.testcase):
    if (identification == 'url'):
        urlName = name
        urlTrue = True
    if (identification == 'sideBar'):
        sideBarName = name
        sideBarTrue = True
    if (identification == 'column'):
        worksheet.write(row, 0, 'Check ' + name + ' column')
        if (urlTrue):
            enterUrl(urlName, row, col)
            row += 1
        if (sideBarTrue):
            sideBar(sideBarName, row, col)
            row += 1
        sortColumn(name, row, col)
        row += 1

workbook.close()