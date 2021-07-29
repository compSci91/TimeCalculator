from datetime import datetime, date, time
import re
import openpyxl


def convert_time_string(time_string):
    time_array = re.split('[,:]', time_string)
    return time(int(time_array[0]), int(time_array[1]), int(time_array[2]), int(time_array[3])*100000)

def calculate_time_difference(start_time_string, end_time_string):
    start_time = convert_time_string(start_time_string)
    end_time = convert_time_string(end_time_string)
    return datetime.combine(date.min, end_time) - datetime.combine(date.min, start_time)


# print(calculate_time_difference("0:00:09,4", "0:00:10,7"))
wb = openpyxl.load_workbook("Amanda.xlsx")

sheet = wb["Sheet1"]
start_time_string = sheet.cell(2,2).value
end_time_string = sheet.cell(2,3).value
time_difference = str(calculate_time_difference(start_time_string, end_time_string)).replace(".", ",", 1)

sheet.cell(2, 4, value=time_difference)
wb.save("Amanda.xlsx")
