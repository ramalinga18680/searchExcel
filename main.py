import sys

sys.path.append('/usr/lib/python3/dist-packages')
import openpyxl


def find_specific_cell(lookFor):
    for row in range(1, currentSheet.max_row + 1):
        for column in "ABCDEFGHIJKLMNOPQRST":  # Here you can add or reduce the columns
            cell_name = "{}{}".format(column, row)
            if currentSheet[cell_name].value == lookFor:
                # print("{1} cell is located on {0}" .format(cell_name, currentSheet[cell_name].value))
                print("cell position {} has value {}".format(cell_name, currentSheet[cell_name].value))
                return cell_name
    return None


def get_column_letter(specificCellLetter):
    letter = specificCellLetter[0:-1]
    print(letter)
    return letter


def get_all_values_by_cell_letter(letter):
    for row in range(1, currentSheet.max_row + 1):
        for column in letter:
            cell_name = "{}{}".format(column, row)
            # print(cell_name)
            print("cell position {} has value {}".format(cell_name, currentSheet[cell_name].value))


def get_search_item(key, value):
    for row in range(1, currentSheet.max_row + 1):
        for column in key:
            cell_name = "{}{}".format(column, row)
            cell_name2 = "{}{}".format(value, row)
            # print(cell_name)
            print(
                "cell position {} has value {} second row value is {}".format(cell_name, currentSheet[cell_name].value,
                                                                              currentSheet[cell_name2].value))


def search_sheet(lookingfor):
    print("searching for", lookingfor)
    for row in range(1, currentSheet.max_row + 1):
        for column in range(1, currentSheet.max_column + 1):  # Here you can add or reduce the columns
            cell_name = "{}{}".format(column, row)
            if lookingfor == currentSheet.cell(row, column).value:
                # if currentSheet[cell_name].value == lookFor:
                # print("{1} cell is located on {0}" .format(cell_name, currentSheet[cell_name].value))
                print("cell found  at row{} and column {}".format(row, column))
                return row, column

    return None, None


def search_row(lookingfor, rownum):
    print("searching for", lookingfor)
    for column in range(1, currentSheet.max_column + 1):
        cell_name = "{}{}".format(column, rownum)
        if lookingfor == currentSheet.cell(rownum, column).value:
            print("cell found  at row{} and column {}".format(rownum, column))
            return rownum, column
    return None, None


def search_column(lookingfor, colnum):
    # print("searching for", lookingfor)
    for row in range(1, currentSheet.max_row + 1):
        cell_name = "{}{}".format(colnum, row)
        if lookingfor == currentSheet.cell(row, colnum).value:
            # print("cell found  at row{} and column {}".format(row, colnum))
            return row, colnum
    return None, None


def load_source_sheet(source_file):
    etilist = []
    srcActiveSheet = source_file[sys.argv[8]]
    collookup = openpyxl.utils.column_index_from_string('I');
    # start from range 2 to avoid taking the header row
    for row in range(2, srcActiveSheet.max_row + 1):
        valueincell = srcActiveSheet.cell(row, collookup).value;
        if valueincell is not None:
            etilist.append(valueincell)

    return etilist;


def writerow(source_file, sheet_to_write, data, row):
    i = 0
    for col in range(2, 5):
        sheet_to_write.cell(row, col).value = data[i]
        i += 1
    source_file.save(sys.argv[7])


# Press the green button in the gutter to run the script.
# argv[1] - excel sheet path and name
# argv[2] - sheet to look in the excel sheet
# argv[3] - item to search for in  the sheet
# argv[4] - search the item in the column
# argv[5],argv[6] row to print when the item searched item is found
# argv[7] - source excel sheet to pick the ETI ID's
# argv[8] Sheet to look from the source workbook
# usage python3 main.py /home/venkata/PycharmProjects/SearchExcel/testdump.xlsx Delivered "ETI-T100453" M H N
# python3 main.py /home/venkata/PycharmProjects/SearchExcel/testdump.xlsx Delivered "ETI-T100453" M H N  /home/venkata/PycharmProjects/SearchExcel/test.xlsx Automate

if __name__ == '__main__':
    print("File to open is {}".format(sys.argv[1]))
    print("Sheet  to open is {}".format(sys.argv[2]))
    print("Item to look for {}".format(sys.argv[3]))
    dstrow: int = 2
    # open the file from where the search has to be done
    theFile = openpyxl.load_workbook(sys.argv[1])
    allSheetNames = theFile.sheetnames
    # open the sheet in which to search
    currentSheet = theFile[sys.argv[2]]
    print("Current sheet name is {} ,max rows is {} max columns is {}".format(currentSheet, currentSheet.max_row,
                                                                              currentSheet.max_column))
    # open the file from where the input areto be taken
    srcFile = openpyxl.load_workbook(sys.argv[7])
    # parse the file and get the data to be searched as a list
    srclist = load_source_sheet(srcFile)
    print("srclist length is ", len(srclist))

    searchin = openpyxl.utils.column_index_from_string(sys.argv[4])

    # create a sheet at index 0 to write the search results
    analysis_sheet = srcFile.create_sheet(title='Analysis', index=0)
    # iterate over the srclist and start searching
    for index in srclist:
        # for index in issuelist:
        row, column = search_column(index, searchin)
        # print("row is {},column is {}".format(row, column))
        if row is not None and column is not None:
            col_to_look1 = openpyxl.utils.column_index_from_string(sys.argv[5])
            col_to_look2 = openpyxl.utils.column_index_from_string(sys.argv[6])
            cell_to_print1 = currentSheet.cell(row, col_to_look1)
            cell_to_print2 = currentSheet.cell(row, col_to_look2)
            towrite = (index, cell_to_print1.value, cell_to_print2.value)
            writerow(srcFile, analysis_sheet, towrite, dstrow)
            dstrow += 1
        else:
            towrite = (index, "TEST_CASE_NOT_FOUND", "CANNOT_DETERMINE_AUTOMATION_STATUS")
            writerow(srcFile, analysis_sheet, towrite, dstrow)
            dstrow += 1

# See PyCharm help at https://www.jetbrains.com/help/pycharm/
