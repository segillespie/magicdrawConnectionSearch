import os
os.chdir("c:/Users/stephen.gillespie/Desktop/Mac/")
from support_functions import excel_file_difference_JSON
from support_functions import write_changes

directory = "c:/Users/stephen.gillespie/Desktop/Mac/"
original_file = 'TALOS CI Table 1650_Baseline.xlsm'
change_files_list = ['TALOS CI Table 1650_Change.xlsm', 'TALOS CI Table 1650_Change2.xlsm']
change_authors_list = ['Gillespie', 'Steve']
date = '25JUL2018'
tabletype = 'CI'
sheet = 'Report'

[JSON_list, no_auto_change, set_of_conflicting_changes, change_list] = excel_file_difference_JSON(directory, original_file, change_files_list, change_authors_list, date, tabletype, sheet)

write_changes(JSON_list, no_auto_change, set_of_conflicting_changes, change_list, directory, tabletype + '_change_' + date + '.xlsx', tabletype + '_change_' + date + '.json')