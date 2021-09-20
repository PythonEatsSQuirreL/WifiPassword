import subprocess

import re

import xlsxwriter

workbook = xlsxwriter.Workbook('Wifipass.xlsx')
worksheet = workbook.add_worksheet()
rowx = 0
col = 0
a = "password"
b = "group"
c = "ssid"
worksheet.write(rowx, col, str(c))
worksheet.write(rowx, col + 1, str(a))
worksheet.write(rowx, col + 2, str(b))
rowx = 1
rowy = 1
col = 1
colx = 1

command_output = subprocess.run(["netsh", "wlan", "show", "profiles"], capture_output = True).stdout.decode()

profile_names = (re.findall("All User Profile     : (.*)\r", command_output))

wifi_list = list()

if len(profile_names) != 0:
    for name in profile_names:
        wifi_profile = dict()
        profile_info = subprocess.run(["netsh", "wlan", "show", "profile", name], capture_output = True).stdout.decode()
        if re.search("Security key            : Absent", profile_info):
            continue
        else:
            wifi_profile["ssid"] = name
            worksheet.write(rowx, col-1, str(name))
            profile_info_pass = subprocess.run(["netsh", "wlan", "show", "profile", name, "key=clear"], capture_output = True).stdout.decode()
            password = re.search("Key Content            : (.*)\r", profile_info_pass)
            if password == None:
                wifi_profile["password"] = None
            else:
                wifi_profile["password"] = password[1]
                worksheet.write(rowx, col, str(password[1]))
                rowx += 1
            wifi_list.append(wifi_profile)

for x in range(len(wifi_list)):
    worksheet.write(rowy, col + 1, str(wifi_list[x]))
    print(wifi_list[x])
    rowy += 1
workbook.close()
