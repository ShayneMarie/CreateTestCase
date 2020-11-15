import xlsxwriter
import testData

workbook = xlsxwriter.Workbook('ClickButtonFunction.xlsx')
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

def textFieldStep(name, row, col):
    worksheet.write(row, col + 1, 'Enter ' + name + ' text')
    worksheet.write(row, col + 2, name + ' should be accepted')

def dropdownTextBox(name, row, col):
    worksheet.write(row, col + 1, 'Select from ' + name + ' list')
    worksheet.write(row, col + 2, 'Selection should be displayed in ' + name + ' dropdown textbox')

def buttonClickToPageStep(name, row, col, error = False):
    worksheet.write(row, col + 1, 'Click ' + name + ' button')
    if error:
        worksheet.write(row, col + 2, '<Error> error message should be displayed')
    else:
        worksheet.write(row, col + 2, 'User should redirect to ' + testData.endPage)

def buttonClickStep(name, row, col, error = False):
    worksheet.write(row, col + 1, 'Click ' + name + ' button')
    if error:
        worksheet.write(row, col + 2, '<Error> error message should be displayed')
    else:
        worksheet.write(row, col + 2, 'Users should be ' + testData.expectedAction)

def successfulClick(test, row):
    worksheet.write(row, 0, 'Successful ' + test)

def unsuccessfulClick(test, field, row, testItem):
    if (testItem == 'empty'):
        worksheet.write(row, 0, 'Unsuccessful ' + test + ' with empty ' + field)
    elif (testItem == 'invalid'):
        worksheet.write(row, 0, 'Unsuccessful ' + test + ' with invalid ' + field)

tupleLen = len(testData.testcase)

row = 1
col = 0

#Check each field
for name, identification in (testData.testcase):
    if (identification == 'url'):
        urlName = name
        urlTrue = True
    if (identification == 'textbox'):
        worksheet.write(row, 0, 'Check ' + name + ' textbox')
        if (urlTrue):
            enterUrl(urlName, row, col)
            row += 1
        textFieldStep(name, row, col)
        row += 1
    elif (identification == 'dropdownTextBox'):
        worksheet.write(row, 0, 'Check ' + name + ' dropdown textbox')
        if (urlTrue):
            enterUrl(urlName, row, col)
            row += 1
        dropdownTextBox(name, row, col)
        row += 1

#Successful
for name, identification in (testData.testcase):
    if (identification == 'testcase'):
        successfulClick(name, row)
    elif (identification == 'url'):
        enterUrl(name, row, col)
        row += 1
    elif (identification == 'textbox'):
        textFieldStep(name, row, col)
        row += 1
    elif (identification == 'dropdownTextBox'):
        dropdownTextBox(name, row, col)
        row += 1
    elif (identification == 'sideBar'):
        sideBar(name, row, col)
        row += 1
    elif (identification == 'buttonClickToPage'):
        buttonClickToPageStep(name, row, col)
        row += 1
    elif (identification == 'buttonClick'):
        buttonClickStep(name, row, col)
        row += 1
    else:
        print('Identification not defined')

#Unsuccessful empty
index = 2
while (index > 1 and index < tupleLen - 1):
    fieldToBeTested = list(testData.testcase[index])
    listTestData = list(testData.testcase)
    listTestData.remove(listTestData[index])
    tupleTestData = tuple(listTestData)
    index += 1

    for name, identification in (tupleTestData):
        if (identification == 'testcase'):
            unsuccessfulClick(name, fieldToBeTested[0], row, 'empty')
        elif (identification == 'url'):
            enterUrl(name, row, col)
            row += 1
        elif (identification == 'textbox'):
            textFieldStep(name, row, col)
            row += 1
        elif (identification == 'sideBar'):
            sideBar(name, row, col)
            row += 1
        elif (identification == 'dropdownTextBox'):
            dropdownTextBox(name, row, col)
            row += 1
        elif (identification == 'buttonClickToPage'):
            buttonClickToPageStep(name, row, col, True)
            row += 1
        elif (identification == 'buttonClick'):
            buttonClickStep(name, row, col, True)
            row += 1
        else:
            print('Identification not defined')

#Unsuccessful invalid
index = 2
while (index > 1 and index < tupleLen - 1):
    fieldToBeTested = list(testData.testcase[index])
    index += 1

    for name, identification in (testData.testcase):
        if (identification == 'testcase'):
            unsuccessfulClick(name, fieldToBeTested[0], row, 'invalid')
        elif (identification == 'url'):
            enterUrl(name, row, col)
            row += 1
        elif (identification == 'sideBar'):
            sideBar(name, row, col)
            row += 1
        elif (identification == 'textbox'):
            if (name == fieldToBeTested[0]):
                textFieldStep('invalid ' + name, row, col)
            else:
                textFieldStep(name, row, col)
            row += 1
        elif (identification == 'buttonClickToPage'):
            buttonClickToPageStep(name, row, col, True)
            row += 1
        elif (identification == 'buttonClick'):
            buttonClickStep(name, row, col, True)
            row += 1
        else:
            print('Identification not defined')

workbook.close()