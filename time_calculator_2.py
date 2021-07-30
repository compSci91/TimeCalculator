import openpyxl

wb = openpyxl.load_workbook("Trelle Dandridge.xlsx")
sheet = wb["Sheet1"]

max_number_of_rows = sheet.max_row

# sheet.cell(row, 1).value

row = 3

while row <= 150:
    # print("Inside while loop")
    print(sheet.cell(row,4).value)
    if(sheet.cell(row, 4).value):
        # print("Starting search")
        start_row = row
        end_row = row + 1

        should_continue = True

        while should_continue:
            if(sheet.cell(end_row, 4).value):
                end_row = end_row + 1
                print("Keep incrementing end_row: ", end_row)
            else:
                print(end_row-1)
                start_time = float(sheet.cell(start_row, 3).value)
                end_time = float(sheet.cell(end_row - 1, 3).value)

                print(end_time - start_time)
                should_continue = False
                row = end_row
    else:
        row = row + 1
