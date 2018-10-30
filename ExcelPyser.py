

#Program: ExcelPyser
#Author: Xdude736
#Purpose: Just a parser I wrote to go through and extract data I needed from some
#   Excel files. Uses the xlrd module.
#Algorithm: This is how the program worksheet.
#   1.) Go through each folder and subfolder until you get to an excel file, Spam
#   2.) Create a new excel file, Spam_mod
#   3.) For each sheet in Spam, create a new sheet in Spam_mod
#   4.) Go through the columns and rows of Spam and place only the data we need in Spam_mod
#   5.) Take Spam_mod and add it to the csv file 'data.csv' as a new row entry for each sheet
#   6.) Rinse and repeat until done

#----------------------------------------------SETUP FOR THE PROGRAM-------------------------------------------------#
import xlrd #for reading excel files
import xlwt #for writing excel files
import csv #for making csv files
import os #for recursivly moving through directories
import datetime #for getting dates into the proper format

#setup the variables
rootdir = "E:\\411_Shtuff\S&W_Invoice_Data\\"
xlWriteRow = 1 #will be incremented to keep from having row collision
eCount = 0
nCount = 0

#formats the date
date_format = xlwt.XFStyle()
date_format.num_format_str = 'mm/dd/yyyy'

#create the excel sheet to write to and setup the header row
writebook = xlwt.Workbook()
writesheet = writebook.add_sheet('Data')
writesheet.write(0,0, 'DATE')
writesheet.write(0,1, 'VIN #')
writesheet.write(0,2, 'MILEAGE')
writesheet.write(0,3, 'VEHICLE')
writesheet.write(0,4, 'PHONE #')
writesheet.write(0,5, 'CUSTOMER')
writesheet.write(0,6, 'ADDRESS')
writesheet.write(0,7, 'DESCRIPTION/AMOUNT')
writesheet.write(0,8, 'SUBTOTAL')
writesheet.write(0,9, 'TAX RATE')
writesheet.write(0,10, 'SALES TAX')
writesheet.write(0,11, 'OTHER')
writesheet.write(0,12, 'TOTAL')
writesheet.write(0,13, 'FILE')

#Method: parseWorkbook
#Purpose: Opens a specific excel file, pulls out specific data from it,
#   places that data in a new excel in a specific format, and then turns that file into a csv
def parseWorkbook(currentFile):
    global xlWriteRow
    global eCount
    #Open that file
    try:
        #Open that file
        workbook = xlrd.open_workbook(currentFile)
        if workbook != None:
            #print(currentFile + ' Opened')

            #go through each sheet in the workbook
            for shidx in range(0, workbook.nsheets):
                xlWriteCol = 0
                #open the worksheet
                worksheet = workbook.sheet_by_index(shidx)

                #check if it's open
                if worksheet == None:
                    print('Worksheet Not Opened')
                #else:
                    #print(worksheet.name + ' Opened')

                #Calcualte the bounds
                if worksheet.cell(rowx=0,colx=4).value != xlrd.empty_cell.value:
                    car_info_bound_low = 1
                    car_info_bound_high = 6
                    cust_info_bound_low = 6
                    cust_info_bound_high = 9
                    work_info_bound_low = 9
                    work_info_bound_high = worksheet.nrows-5
                    value_info_bound_low = worksheet.nrows-5
                    value_info_bound_high = worksheet.nrows
                    #print('test test test')
                else:
                    car_info_bound_low = 1 + 1
                    car_info_bound_high = 6 + 1
                    cust_info_bound_low = 6 + 1
                    cust_info_bound_high = 9 + 1
                    work_info_bound_low = 9 + 1
                    work_info_bound_high = worksheet.nrows-5
                    value_info_bound_low = worksheet.nrows-5
                    value_info_bound_high = worksheet.nrows

                #Get the section of data from rows 1-6, range does not include the lower bound in python
                #colx is set to 4 because that is the column where our data is consistently
                for row_index in range(car_info_bound_low, car_info_bound_high):
                    if worksheet.cell(rowx=row_index,colx=4).value == xlrd.empty_cell.value:
                        xlWriteCol += 1
                    elif worksheet.cell(rowx=row_index,colx=4).ctype == xlrd.XL_CELL_DATE:
                        dt_tuple = xlrd.xldate_as_tuple(worksheet.cell(rowx=row_index,colx=4).value, workbook.datemode)
                         # Create datetime object from this tuple.
                        cellValue = datetime.datetime(*dt_tuple)
                        #print(cellValue)
                        writesheet.write(xlWriteRow, xlWriteCol, cellValue, date_format) #writes on the next row of our workbook, pulls from the current row of the client data
                        xlWriteCol += 1
                    else:
                        cellValue = worksheet.cell(rowx=row_index,colx=4).value
                        writesheet.write(xlWriteRow, xlWriteCol, cellValue) #writes on the next row of our workbook, pulls from the current row of the client data
                        xlWriteCol += 1
                        #print(cellValue)


                #get the customer billing data in row 7-9
                cellValue = worksheet.cell(rowx=cust_info_bound_low,colx=4).value
                writesheet.write(xlWriteRow, xlWriteCol, cellValue)
                xlWriteCol += 1
                cust_info = ''
                for row_index in range(cust_info_bound_low+1, cust_info_bound_high):
                    if worksheet.cell(rowx=row_index,colx=4).value != xlrd.empty_cell.value:
                        cellValue = worksheet.cell(rowx=row_index,colx=4).value
                        if row_index == 8:
                            cust_info += str(cellValue)
                        else:
                            cust_info += str(cellValue) + '\n'
                writesheet.write(xlWriteRow, xlWriteCol, cust_info)
                xlWriteCol += 1

                #get the descriptions and the amounts
                description = ''
                for row_index in range(work_info_bound_low, work_info_bound_high):
                    for col_index in range(0, 5):
                        if worksheet.cell(rowx=row_index,colx=col_index).value != xlrd.empty_cell.value:
                            cellValue = worksheet.cell(rowx=row_index,colx=col_index).value
                            description += str(cellValue) + ';'
                writesheet.write(xlWriteRow, xlWriteCol, description)
                xlWriteCol += 1

                #Start at column 8 for writing to the excel sheet
                col_num = 8
                for row_index in range(value_info_bound_low, value_info_bound_high):
                    #if worksheet.cell(rowx=row_index,colx=4).value != xlrd.empty_cell.value:
                    cellValue = worksheet.cell(rowx=row_index,colx=4).value
                    writesheet.write(xlWriteRow, col_num, cellValue)
                    col_num += 1

                #puts in the file path in the last column, used for easy lookup
                writesheet.write(xlWriteRow, col_num, currentFile)
                xlWriteRow += 1
        else:
            print('File Not Opened')
    except Exception as e:
        error = str(e)
        eCount += 1
        print(os.path.join(root, filename))
        print('Error with file: ' + error)
        #pass


#-----------------------------------------------------START MAIN PROGRAM-------------------------------------------------#

#recusively move through the directories
for root, subFolders, files in os.walk(rootdir):
    #print('-----------------------------------------------')
    for filename in files:
        #print(os.path.join(root, filename))
        name = str(filename)
        #Do a check to see if it is an excel file, if it is do the thang, if not log the filename
        if name.lower().endswith('.xlsx'):
            parseWorkbook(os.path.join(root, filename))
        else:
            #print(os.path.join(root, filename))
            #print('Narp')
            nCount += 1

#display the output at the end
writebook.save('Legacy_Data.xls')
print(eCount)
print(nCount)
