#!/usr/bin/python

import pandas
import numpy # used but not directly referenced
import csv
import json
import xlrd # flagged here just so pipreqs sees it

# target must be resaved as a real .xls file, not the half xml Section Tally produces
# actual format is .xml but the parser find a bug somewhere
display_start_time = ['0800', '0930', '1100', '1230', '1400', '1530', '1700', '1830', '2000']
display_end_time = ['0915', '1045', '1215', '1345', '1515', '1645', '1815', '1945', '2115']

def parse_section_tally(target):
    try:
        df = pandas.read_excel(target, usecols='A,C:D,G:I', engine='xlrd')
        # df = pandas.read_xml(target)
    except Exception as e:
        print("Tip: Section Tally files must be resaved as true .xls first.")
        print(e)
        exit(1)
        
    room_col = df.columns[-1] # former col I
    # dupe data (find \n in col I and make new entry)
    df1 = df[room_col].str.split('\n', expand=True)\
        .stack()\
        .reset_index(level=1)\
        .rename(columns={0: room_col})\
        .drop(axis=1, labels='level_1')
    # merge back in
    df = df.drop(axis=1, labels=room_col)\
        .merge(df1, left_index=True, right_index=True)\
        .reset_index(drop=True)
    # split room_col to individual col (split is strange on classes without rooms due to .split())
    df2 = df[room_col].str.split(expand=True)\
        .set_axis(room_col.split(), axis=1)
    # merge back in
    df = df.merge(df2, left_index=True, right_index=True)\
        .drop(axis=1, labels=room_col)

    # filter to only Rowan Hall or Eng. Hall. This has the side effect of dropping Online courses or courses with unmarked rooms
    df = df[df['Bldg'].isin(['ROWAN', 'ENGR'])]
    # handle multiple profs, '\n' to '; \n'
    df['Prof'] = df['Prof'].str.rstrip().replace('\n','; \n', regex=True)
    return df

def save_to_excel(dataframe, filename):
    with pandas.ExcelWriter(filename) as writer:
        dataframe.to_excel(writer, sheet_name='parsed', index=False)

def map_course_names(df, map_dict):
    df['Title'] = df['Title'].replace(map_dict)
    return df

def room_occupancy(df, prof =  '.', building = '.', room = '.', day = '.'):
    _df = df[df['Prof'].str.contains(prof) \
            & df['Day'].str.contains(day) \
            & df['Bldg'].str.contains(building) \
            & df['Room'].str.contains(room)]
    return _df.sort_values(by=['Room', 'Beg'])

def room_occupancy_on_day(_df, _building, _room, _day):
    __df = room_occupancy(_df, building=_building, room=_room, day=_day)
    __df = __df.loc[:,['Beg','Title','Prof']]
    __df = __df.set_index('Beg').reindex(index=display_start_time).reset_index().fillna(value="")
    # print(_df)
    __df = __df.to_numpy()
    return [_day, _building + _room], __df
    # for i, time in enumerate(display_start_time):
    #     if _df[i][0]

def pretty_print(df, _bldg, _rooms, days = ['M', 'T', 'W', 'T', 'F']):
    display_array = [[""],[""]]
    num_col = 2 # adjust if more than instructor and class are needed
    first_pass = True
    for day in days:
        display_array[0].append(day)
        for i in range(num_col*len(rooms) -1):
            display_array[0].append("")
        for room in rooms:
            display_array[1].append(bldg + room)
            for i in range(num_col -1):
                display_array[1].append("")
            key, array = room_occupancy_on_day(df,  bldg, room, day)
            if first_pass:
                first_pass = False
                display_array.extend(array.tolist())
            else:
                for i, row in enumerate(array[:,1:].tolist()):
                    display_array[i+2].extend(row)
    return display_array

if __name__ == '__main__':
    """ IMPORTANT
    whatever is downloaded from Section Tally MUST be resaved as a true .xls
    Section Tally outputs a broken xml file as far as I can tell
    """
    section_tally_target = 'section_tally_f23_resave.xls' 
    section_tally_output = 'section_tally_f23_parsed.xlsx'
    bldg = 'ENGR' # currently supports ENGR and ROWAN, case dependant
    rooms = ['140', '141', '240', '241'] # must be a list, even if single entry
    with open('course_title_dict.json', 'r') as f:
        course_title_dict = json.load(f)
    df = parse_section_tally(section_tally_target)
    df = map_course_names(df, course_title_dict) # exact names only for now
    save_intermediate = False
    if save_intermediate:
        save_to_excel(df, section_tally_output)
    pretty_array = pretty_print(df, bldg, rooms)
    with open('test.csv', 'w', newline='') as f:
        w = csv.writer(f)
        w.writerows(pretty_array)
    