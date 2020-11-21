import xlsxwriter
from testData import testDataClick

workbook = xlsxwriter.Workbook('ExpectedResult/ClickButtonFunction.xlsx')
worksheet = workbook.add_worksheet()

worksheet.write(0, 0, 'Test Case')
worksheet.write(0, 1, 'Steps')
worksheet.write(0, 2, 'Expected Result')


def enterUrl(name, expected, row, col):
    worksheet.write(row, col + 1, f'Go to {name}')
    worksheet.write(row, col + 2, f"{expected}")


def sideBar(name, row, col):
    worksheet.write(row, col + 1, f'Click {name} sidebar')
    worksheet.write(row, col + 2, f'{name} page should be displayed')


def link(name, row, col):
    worksheet.write(row, col + 1, f'Click {name} link')
    worksheet.write(row, col + 2, f'{name} page should be displayed')


def textFieldStep(name, row, col):
    worksheet.write(row, col + 1, f'Enter {name} text')
    worksheet.write(row, col + 2, f'{name} should be accepted')


def dropdownTextBox(name, row, col):
    worksheet.write(row, col + 1, f'Select from {name} list')
    worksheet.write(row, col + 2, f'Selection should be displayed in {name} dropdown textbox')


def buttonClick(name, expected, row, col, error=False):
    worksheet.write(row, col + 1, f'Click {name} button')
    if error:
        worksheet.write(row, col + 2, '<Error> error message should be displayed')
    else:
        worksheet.write(row, col + 2, f'{expected}')


def validLogin(row, col):
    worksheet.write(row, col + 1, f'Enter valid credentials')
    worksheet.write(row, col + 2, f'{testDataClick.mainPage} should be displayed')


def successfulClick(test, row):
    worksheet.write(row, 0, 'Successful ' + test)


def unsuccessfulClick(test, field, row, testItem):
    if testItem == 'empty':
        worksheet.write(row, 0, f'Unsuccessful {test} with empty ' + field)
    elif testItem == 'invalid':
        worksheet.write(row, 0, f'Unsuccessful {test} with invalid {field}')


tupleLen = len(testDataClick.testcase)

row = 1
col = 0

# Check each field
for name, identification, expected in testDataClick.testcase:
    if identification == 'textbox':
        worksheet.write(row, 0, f'Check {name} textbox')
        #counter = 1
        for nname, identification, expected in testDataClick.testcase:
            #if counter == tupleLen:
             #   break
            if identification == 'url':
                enterUrl(nname, expected, row, col)
            elif identification == 'link':
                link(nname, row, col)
            elif identification == 'login':
                validLogin(row, col)
            elif identification == 'sideBar':
                sideBar(nname, row, col)
            elif identification == 'buttonClick':
                buttonClick(nname, expected, row, col)
            else:
                row -= 1
            row += 1
            #counter += 1
        textFieldStep(name, row, col)
        row += 1
    elif identification == 'dropdownTextBox':
        worksheet.write(row, 0, f'Check {name} dropdown textbox')
        counter = 1
        for nname, identification, expected in testDataClick.testcase:
            #if counter == tupleLen:
            #    break
            if identification == 'url':
                enterUrl(nname, expected, row, col)
            elif identification == 'link':
                link(nname, row, col)
            elif identification == 'login':
                validLogin(row, col)
            elif identification == 'sideBar':
                sideBar(nname, row, col)
            elif identification == 'buttonClick':
                buttonClick(nname, expected, row, col)
            else:
                row -= 1
            row += 1
            #counter += 1
        dropdownTextBox(name, row, col)
        row += 1

# Successful
for name, identification, expected in testDataClick.testcase:
    if identification == 'testcase':
        successfulClick(name, row)
    elif identification == 'url':
        enterUrl(name, expected, row, col)
        row += 1
    elif identification == 'link':
        link(name, row, col)
        row += 1
    elif identification == 'login':
        validLogin(row, col)
        row += 1
    elif identification == 'textbox':
        textFieldStep(name, row, col)
        row += 1
    elif identification == 'dropdownTextBox':
        dropdownTextBox(name, row, col)
        row += 1
    elif identification == 'sideBar':
        sideBar(name, row, col)
        row += 1
    elif identification == 'buttonClick':
        buttonClick(name, expected, row, col)
        row += 1
    else:
        print('Identification not defined')

# Unsuccessful empty
index = 2
while index < tupleLen - 1:
    fieldToBeTested = list(testDataClick.testcase[index])
    listTestData = list(testDataClick.testcase)
    if fieldToBeTested[1] == 'textbox' or fieldToBeTested[1] == 'dropdownTextBox':
        listTestData.remove(listTestData[index])
    tupleTestData = tuple(listTestData)
    index += 1
    counter = 1 # for multiple button check

    for name, identification, expected in (tupleTestData):
        if fieldToBeTested[1] == 'login' or fieldToBeTested[1] == 'buttonClick' or fieldToBeTested[1] == 'sideBar' or \
                fieldToBeTested[1] == 'link':
            break
        if identification == 'testcase':
            unsuccessfulClick(name, fieldToBeTested[0], row, 'empty')
        elif identification == 'url':
            enterUrl(name, expected, row, col)
            row += 1
        elif identification == 'link':
            link(name, row, col)
            row += 1
        elif identification == 'login':
            validLogin(row, col)
            row += 1
        elif identification == 'textbox':
            textFieldStep(name, row, col)
            row += 1
        elif identification == 'sideBar':
            sideBar(name, row, col)
            row += 1
        elif identification == 'dropdownTextBox':
            dropdownTextBox(name, row, col)
            row += 1
        elif identification == 'buttonClick':
            if counter == tupleLen-1:
                buttonClick(name, expected, row, col, True)
            else:
                buttonClick(name, expected, row, col)
            row += 1
        else:
            print('Identification not defined')
        counter += 1

# Unsuccessful invalid
index = 2
while index > 1 and index < tupleLen - 1:
    fieldToBeTested = list(testDataClick.testcase[index])
    index += 1
    counter = 1

    for name, identification, expected in testDataClick.testcase:
        if fieldToBeTested[1] == 'login' or fieldToBeTested[1] == 'buttonClick' or fieldToBeTested[1] == 'sideBar' or \
                fieldToBeTested[1] == 'link':
            break
        if identification == 'testcase':
            unsuccessfulClick(name, fieldToBeTested[0], row, 'invalid')
        elif identification == 'url':
            enterUrl(name, expected, row, col)
            row += 1
        elif identification == 'link':
            link(name, row, col)
            row += 1
        elif identification == 'login':
            validLogin(row, col)
            row += 1
        elif identification == 'sideBar':
            sideBar(name, row, col)
            row += 1
        elif identification == 'textbox':
            if name == fieldToBeTested[0]:
                textFieldStep('invalid ' + name, row, col)
            else:
                textFieldStep(name, row, col)
            row += 1
        elif identification == 'buttonClick':
            if counter == tupleLen:
                buttonClick(name, expected, row, col, True)
            else:
                buttonClick(name, expected, row, col)
            row += 1
        else:
            print('Identification not defined')
        counter += 1

workbook.close()
