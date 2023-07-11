import numpy as np
import pandas as pd
import os
import time
us_state_to_abbrev = {
    "Alabama": "AL",
    "Alaska": "AK",
    "Arizona": "AZ",
    "Arkansas": "AR",
    "California": "CA",
    "Colorado": "CO",
    "Connecticut": "CT",
    "Delaware": "DE",
    "Florida": "FL",
    "Georgia": "GA",
    "Hawaii": "HI",
    "Idaho": "ID",
    "Illinois": "IL",
    "Indiana": "IN",
    "Iowa": "IA",
    "Kansas": "KS",
    "Kentucky": "KY",
    "Louisiana": "LA",
    "Maine": "ME",
    "Maryland": "MD",
    "Massachusetts": "MA",
    "Michigan": "MI",
    "Minnesota": "MN",
    "Mississippi": "MS",
    "Missouri": "MO",
    "Montana": "MT",
    "Nebraska": "NE",
    "Nevada": "NV",
    "New Hampshire": "NH",
    "New Jersey": "NJ",
    "New Mexico": "NM",
    "New York": "NY",
    "North Carolina": "NC",
    "North Dakota": "ND",
    "Ohio": "OH",
    "Oklahoma": "OK",
    "Oregon": "OR",
    "Pennsylvania": "PA",
    "Rhode Island": "RI",
    "South Carolina": "SC",
    "South Dakota": "SD",
    "Tennessee": "TN",
    "Texas": "TX",
    "Utah": "UT",
    "Vermont": "VT",
    "Virginia": "VA",
    "Washington": "WA",
    "West Virginia": "WV",
    "Wisconsin": "WI",
    "Wyoming": "WY",
    "District of Columbia": "DC",
    "American Samoa": "AS",
    "Guam": "GU",
    "Northern Mariana Islands": "MP",
    "Puerto Rico": "PR",
    "United States Minor Outlying Islands": "UM",
    "U.S. Virgin Islands": "VI",
}

folders_path = 'C:/Users/blake/Documents/VSCode/Python/Hope/Cision Files/'
output_path = folders_path

cheat_sheet = np.loadtxt(folders_path+'../'+'FormattingCheatSheet.txt', delimiter='|', skiprows=2,dtype=str)
cheat_dict = {}
for row in cheat_sheet:
    cheat_dict[row[1].strip()] = row[0].strip()
folder_name = input('Input Folder Name: ')
try:
    files = os.listdir(folders_path+folder_name)
except:
    print(folder_name + ' not found :(')
    exit()

if len(files) == 0:
    print('No files in '+folder_name)
    exit()
elif len(files) > 1:
    print('Too many Cision files in '+folder_name+'. Remove all but one file')
    exit()


print('File found!')
input_dat = pd.read_excel(folders_path+folder_name+'/'+files[0])
input_dat = np.array(input_dat)
input_entries = np.transpose(input_dat[4:])
# print(input_entries)
total_hits = int(input_dat[0][1])
cols = input_dat[2]

#input formatting
my_str = 'Default'
output_str = 'Current Format: '
output_cols = []
while True:
    print()
    print(output_str)
    my_str = input('Input Column Heading Shortcut: ').strip()
    if my_str in cheat_dict:
        if cheat_dict[my_str] == 'Finish':
            if my_str == 'I love Blake':
                print('\nAww love you too <3\n')
                time.sleep(2)
            break
        elif cheat_dict[my_str] == 'Blank Column' or cheat_dict[my_str] == 'Topic' or cheat_dict[my_str] == 'Notes':
            output_cols.append('None')
            output_str += cheat_dict[my_str] + '|'
        elif cheat_dict[my_str] == 'Market':
            output_cols.append('M')
            output_str += cheat_dict[my_str] + '|'
        elif cheat_dict[my_str] == 'Delete Previous':
            if len(output_cols) > 0:
                output_cols.pop()
                output_str = output_str[:-1]
                while output_str[-1] != '|':
                    output_str = output_str[:-1]
        else:
            output_cols.append(np.where(cols == cheat_dict[my_str])[0][0])
            output_str += cheat_dict[my_str] + '|'
    else:
        print('Shortcut not found.')

#make output structure
row_count = len(input_entries[0])
col_count = len(output_cols)
output_array = [] #this is the transpose of the true output
for i, col in enumerate(output_cols):
    if col == 'M':
        output_array.append([])
        c = np.where(cols == 'Country')[0][0]
        s = np.where(cols == 'State')[0][0]
        t = np.where(cols == 'City')[0][0]
        for j in range(row_count):
            if input_entries[c][j] != 'United States':
                output_array[i].append(input_entries[c][j])
            else:
                if str(input_entries[t][j]) == 'nan' and str(input_entries[t][j]) == 'nan':
                    output_array[i].append('United States')
                elif str(input_entries[t][j]) == 'nan':
                    output_array[i].append(input_entries[s][j])
                else:
                    output_array[i].append(str(input_entries[t][j])+', '+ us_state_to_abbrev[str(input_entries[s][j])])
    elif col == 'None':
        output_array.append(['']*row_count)
    else:
        output_array.append([])
        for j in range(row_count):
            output_array[i].append(input_entries[col][j])
column_headings = output_str.split('|')
column_headings[0] = column_headings[0][16:]
column_headings.pop()
for i in range(len(column_headings)):
    if column_headings[i] == 'Blank Column':
        column_headings[i] = ''
df = pd.DataFrame(np.transpose(output_array), columns = column_headings)
output_filename = input('Output file name: ')
if output_filename[-5:] == '.xlsx':
    df.to_excel(output_path+output_filename, index=False)
else:
    df.to_excel(output_path+output_filename+'.xlsx', index=False)