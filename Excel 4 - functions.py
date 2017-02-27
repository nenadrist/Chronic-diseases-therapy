# -*- coding: utf-8 -*-
"""
Created on Tue Feb 21 14:53:00 2017

@author: Nenad
"""

# -*- coding: utf-8 -*-
"""
Created on Mon Feb 20 10:17:48 2017

@author: Nenad
"""

# Importing official document from WHO,
# International Disease Classification, ICD-10, but in Serbian translation

import openpyxl
wb = openpyxl.load_workbook('MKB10.xlsx', data_only=True)
sheets = wb.get_sheet_names()

#Look for appropriate worksheet

if len(sheets) > 1:
    for i in range(len(sheets)):
        if wb.get_sheet_by_name(sheets[i]) == '10 revizija bolesti':
            current_sheet = wb.get_sheet_by_name(sheets[i])
        else:
            print (i)
else:
    current_sheet = wb.get_sheet_by_name(sheets[0])

# Picking up proper columns and raws for disease code and
# disease title from determined worksheet - the goal is to get only data
# that will be used in a project, keeping in mind possible changes in file or data structure 

def pick_columns(from_raw, to_raw, from_column, to_column, active_sheet):
    disease = []
    for col in active_sheet.iter_cols(min_row=from_raw, max_col=to_column, max_row=to_raw, min_col=from_column):
        for cell in col:
        #print(cell.value)
            disease.append(cell.internal_value)
    return disease

# Function call and creating lists and corresponding dictionary

disease_code = []
disease_code = (pick_columns(1, 34, 1, 1, current_sheet))
disease_title = []
disease_title = (pick_columns(1, 34, 3, 3, current_sheet))

recnik = dict(zip(disease_code, disease_title))
print ('Rečnik: ', recnik)

# Write dictionary to text file for further use

with open('MKB-dictionary.txt', 'w') as final:
        for k, v in recnik.items():
            line = '{}, {}'.format(k, v) 
            print(line, file=final)
