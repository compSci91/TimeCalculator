from datetime import datetime, date, time
import re
# time1 = time(9, 30, 45, 10000)
# time2 = time(9, 30, 45, 9999)
#
# print(datetime.combine(date.min, time1) - datetime.combine(date.min, time2))

time_string = "0:00:09,4"
# time_array = time_string.split(":,") + [0]
time_array = re.split('[,:]', time_string)
print(time_array)
# seconds_and_milliseconds = time_array[2].split(",")
# print(seconds_and_milliseconds)
# time_array[2] = seconds_and_milliseconds[0]
# time_array[3] = seconds_and_milliseconds[1]
# print(time_array)
# 0:00:10,7
