import xlsxwriter
from testData import testDataColumn

workbook = xlsxwriter.Workbook('ExpectedResult/TableFunction.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, 'Test Case')
worksheet.write(0, 1, 'Steps')
worksheet.write(0, 2, 'Expected Result')

def enterUrl(name, expected, row, col):
    worksheet.write(row, col + 1, f'Login to {name}')
    worksheet.write(row, col + 2, f'{expected}')

def sideBar(name, row, col):
#    worksheet.write(row, col + 1, f'Click {name} sidebar')
#    worksheet.write(row, col + 2, f'{name} page should be displayed')
    worksheet.write(row, col + 1, f'Send automatic {name}')
    worksheet.write(row, col + 2, f'{name} should be received')

def column(name, row, col):
    worksheet.write(row, col + 1, f'Check {name} column')
    worksheet.write(row, col + 2, f'{name} page should be displayed')

def sortColumn(name, row, col):
    worksheet.write(row, col + 1, f'Click {name} column')
    worksheet.write(row, col + 2, f'{name} column should be sorted')

def buttonClick(name, expected, row, col, error = False):
    worksheet.write(row, col + 1, f'Click {name} button')
    if error:
        worksheet.write(row, col + 2, '<Error> error message should be displayed')
    else:
        worksheet.write(row, col + 2, f'{expected}')

row = 1
col = 0

#Check column
for name, identification, expected in (testDataColumn.testcase):
    if identification == 'column':
        worksheet.write(row, 0, f'Check {name} column')
        for nname, identification, expected in (testDataColumn.testcase):
            if identification == 'url':
                enterUrl(nname, expected, row, col)
            elif identification == 'sideBar':
                sideBar(nname, row, col)
            elif identification == 'buttonClick':
                buttonClick(nname, expected, row, col)
            else:
                row -= 1
            row += 1
        column(name, row, col)
        row += 1

#Check column sorting
for name, identification, expected in testDataColumn.testcase:
    if identification == 'column':
        worksheet.write(row, 0, f'Check {name} column')
        for nname, identification, expected in testDataColumn.testcase:
            if identification == 'url':
                enterUrl(nname, expected, row, col)
            elif identification == 'sideBar':
                sideBar(nname, row, col)
            elif identification == 'buttonClick':
                buttonClick(nname, expected, row, col)
            else:
                row -= 1
            row += 1
        sortColumn(name, row, col)
        row += 1

workbook.close()