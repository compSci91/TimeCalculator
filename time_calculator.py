from datetime import datetime, date, time, timedelta
import re
import openpyxl


def convert_time_string(time_string):
    time_array = re.split('[,:]', time_string)
    return time(int(time_array[0]), int(time_array[1]), int(time_array[2]), int(time_array[3])*100000)

def calculate_time_difference(start_time_string, end_time_string):
    start_time = convert_time_string(start_time_string)
    end_time = convert_time_string(end_time_string)
    return datetime.combine(date.min, end_time) - datetime.combine(date.min, start_time)

def convert_duration_string(duration_string):
    time_array = re.split('[,:]', duration_string)
    # return timedelta(int(time_array[0]), int(time_array[1]), int(time_array[2]), int(time_array[3]))
    return timedelta(weeks=0, days=0, hours=int(time_array[0]), minutes=int(time_array[1]), seconds=int(time_array[2]), milliseconds=0, microseconds = int(time_array[3]))


        # return time(int(time_array[0]), int(time_array[1]), int(time_array[2]), int(time_array[3]))



# print(calculate_time_difference("0:00:09,4", "0:00:10,7"))
wb = openpyxl.load_workbook("Amanda.xlsx")

sheet = wb["Sheet1"]

max_number_of_rows = sheet.max_row

# for row in range(2, max_number_of_rows + 1):
#     start_time_string = sheet.cell(row,2).value
#     end_time_string = sheet.cell(row,3).value
#     time_difference = str(calculate_time_difference(start_time_string, end_time_string)).replace(".", ",", 1)
#     sheet.cell(row, 4, value=time_difference)


total_duration = timedelta(weeks=0, days=0, hours=0, minutes=0, seconds=0, milliseconds=0, microseconds=0)
print(total_duration)
print()
for row in range(2, max_number_of_rows + 1):
    if(sheet.cell(row, 1).value == "1. Touching expression"):
        current_duration_string = sheet.cell(row,4).value
        print("Current duration string: ", current_duration_string)
        current_duration = convert_duration_string(current_duration_string)
        # total_duration = (datetime.combine(date.min, current_duration) + total_duration).timedelta
        total_duration = total_duration + current_duration
        print("Total duration string: ", total_duration)
        print()

print(total_duration)




wb.save("Amanda.xlsx")
