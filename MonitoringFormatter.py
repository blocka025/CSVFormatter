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

def get_not_nan_count(row):
    count = 0
    for r in row:
        if str(r) != 'nan':
            count +=1
    return count
    
def get_start(data):
    for i in range(len(data)):
        if get_not_nan_count(data[i]) > 5:
            return i
def get_data_rows(data):
    s = get_start(data)
    output = []
    for i in range(len(data)):
        if i != s:
            if get_not_nan_count(data[i]) > 5:
                output.append(data[i])
    return np.array(output)


folders_path = 'Cision Files/'
output_path = 'Formatted Files/'

cheat_sheet = np.loadtxt('FormattingCheatSheet.txt',
                         delimiter='|',
                         skiprows=2,
                         dtype=str)
cheat_sheet2 = np.loadtxt('NewFormattingCheatSheet.txt',
                         delimiter='|',
                         skiprows=2,
                         dtype=str)
cheat_sheet3 = np.loadtxt('CriticalMentionFormatting.txt',
                         delimiter='|',
                         skiprows=2,
                         dtype=str)
cheat_dict = {}
for row in cheat_sheet:
    cheat_dict[row[1].strip()] = row[0].strip()
for i, row in enumerate(cheat_sheet2):#inport other cheatsheet
    if i < len(cheat_sheet2)-4 and row[1].strip() in cheat_dict and cheat_dict[row[1].strip()] != row[0].strip():
        print(row[1].strip() + ' is being overwritten with '+row[0].strip())
    cheat_dict[row[1].strip()] = row[0].strip()
for i, row in enumerate(cheat_sheet3):#inport third cheatsheet
    if i < len(cheat_sheet3)-4 and row[1].strip() in cheat_dict and cheat_dict[row[1].strip()] != row[0].strip():
        print(row[1].strip() + ' is being overwritten with '+row[0].strip())
    cheat_dict[row[1].strip()] = row[0].strip()
    
folder_name = input('Input Folder Name: ')
try:
    files = os.listdir(folders_path + folder_name)
except:
    print(folder_name + ' not found :(')
    exit()

if len(files) == 0:
    print('No files in ' + folder_name)
    exit()
elif len(files) > 1:
    print('Too many Cision files in ' + folder_name +
          '. Remove all but one file')
    exit()

print('File found!')
path = folders_path + folder_name + '/' + files[0]
if path[-3:] == 'csv':
    input_dat = pd.read_csv(path,header=None)
elif path[-4:] == 'xlsx' or path[-3:] == 'xls':
    input_dat = pd.read_excel(folders_path + folder_name + '/' + files[0])
else:
    print('File type not recognized')
    exit()

input_dat = list(np.array(input_dat))
input_dat2 = get_data_rows(input_dat)
input_entries = np.transpose(input_dat2)
# print(input_entries)
# total_hits = int(input_dat[0][1])
cols = input_dat[get_start(input_dat)]
# print(cols)
for i, col in enumerate(cols):
    # print(col)
    if str(col) =='nan':
        cols[i] = 'nan'
    else:
        cols[i] = col.strip()

while True:
    my_str = input('Show Available Columns? (y/n) : ')
    if my_str == 'y':
        for col in cols:
            print(col)    
        break
    elif my_str == 'n':
        break
    else:
        print('Please type either "y" or "n"')
sort_states = False
while 'State' in cols:
    my_str = input('Sort by northern Midwestern states? (y/n) : ')
    if my_str == 'y':
        sort_states = True   
        break
    elif my_str == 'n':
        break
    else:
        print('Please type either "y" or "n"')


#input formatting
my_str = 'Default'
output_str = 'Current Format: '
output_cols = []
while True:
    print()
    print(output_str)
    my_str = input('Input Column Heading Shortcut: ').strip()
    try:
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
    except IndexError:
        print('The column heading "'+cheat_dict[my_str]+ '" could not be found')

#make output structure
row_count = len(input_entries[0])
col_count = len(output_cols)
output_array = []  #this is the transpose of the true output
for i, col in enumerate(output_cols):
    if col == 'M':
        try:
            output_array.append([])
            c = np.where(cols == 'Country')[0][0]
            s = np.where(cols == 'State')[0][0]
            t = np.where(cols == 'City')[0][0]
            for j in range(row_count):
                if input_entries[c][j] != 'United States':
                    output_array[i].append(input_entries[c][j])
                else:
                    if str(input_entries[t][j]) == 'nan' and str(
                            input_entries[t][j]) == 'nan':
                        output_array[i].append('United States')
                    elif str(input_entries[t][j]) == 'nan':
                        output_array[i].append(input_entries[s][j])
                    else:
                        output_array[i].append(
                            str(input_entries[t][j]) + ', ' +
                            us_state_to_abbrev[str(input_entries[s][j])])
        except:
            output_array.append([])
            for j in range(row_count):
                output_array[i].append(input_entries[col][j])
    elif col == 'None':
        output_array.append([''] * row_count)
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
if sort_states:
    final_out = []
    output_array = np.transpose(output_array)
    s = np.where(cols == 'State')[0][0]
    for i in range(row_count):
        if input_entries[s][i] == 'Minnesota' or input_entries[s][i] == 'Wisconsin' or input_entries[s][i] == 'North Dakota' or input_entries[s][i] == 'South Dakota' or input_entries[s][i] == 'Iowa':
            final_out.append(output_array[i])
    for i in range(row_count):
        if input_entries[s][i] != 'Minnesota' and input_entries[s][i] != 'Wisconsin' and input_entries[s][i] != 'North Dakota' and input_entries[s][i] != 'South Dakota' and input_entries[s][i] != 'Iowa':
            final_out.append(output_array[i])
    output_array = np.transpose(final_out)
df = pd.DataFrame(np.transpose(output_array), columns=column_headings)
output_filename = input('Output file name: ')
df.to_csv(output_path + output_filename+'.csv', index=False)
    

print()
while True:
    my_str = input('Would you like remove old Cision file? (y/n): ')
    if my_str == 'y':
        os.remove(folders_path + folder_name + '/' + files[0])
        break
    elif my_str == 'n':
        break
    else:
        print('Please type either "y" or "n"')
print('Success!! :)')
time.sleep(3)
os.system('clear')
