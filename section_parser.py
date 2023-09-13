#!/usr/bin/python

import pandas
import numpy # used but not directly referenced
import csv
import json
import re
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
    df['Prof'] = df['Prof'].str.rstrip().replace('\n',';', regex=True)
    return df

def save_to_excel(dataframe, filename):
    with pandas.ExcelWriter(filename) as writer:
        dataframe.to_excel(writer, sheet_name='parsed', index=False, header = False)

def map_course_names(df, _dict):
    df['Title'] = df['Title'].replace(_dict)
    return df

def drop_names_not_in(_df, instr_list):
    series_lists = _df['Prof'].str.split(',')
    new_list = []
    for _list in series_lists:
        new_list.append("".join(filter(lambda i: i in instr_list, list(map(str.strip,_list)))))
    _df['Prof'] = new_list
    return _df

def instructor_last_names(_df):
    new_df = _df.filter(['CRN','Prof'], axis=1)
    new_df = new_df['Prof'].str.split(';', expand = True)
    new_df = new_df.replace([
        re.compile(r'\n'),
        re.compile(r' &'),
        re.compile(r'^ ')
         ], '', regex = True).fillna('')
    for colname, coldata in new_df.items():
        temp = coldata.str.split(',', expand = True).drop(columns=[1])
        new_df[colname] = temp.fillna("")
    
    # map to only ENGR faculty last names

    # concat
    # if we find out how to use the string join properly, this can be vastly simplified
    new_df['Prof'] = new_df[0] + ', ' + new_df[1] + ', ' + new_df[2]
    new_df = new_df.drop(columns = [0, 1, 2])
    new_df = new_df.replace([
        re.compile(r'^, , '),
        re.compile(r'^, '),
        re.compile(r', , $'),
        re.compile(r', $')
        ],
        '', regex = True)
    new_df['Prof'] = new_df['Prof'].replace("", "ERROR")
    prof_df = new_df.merge(df.drop(columns = 'Prof'), left_index=True, right_index=True)
    # join dataframe back
    return prof_df

def room_occupancy(df, prof =  '.', room = ('.','.'), day = '.'):
    new_df = df[df['Prof'].str.contains(prof) \
            & df['Day'].str.contains(day) \
            & df['Bldg'].str.contains(room[0]) \
            & df['Room'].str.contains(room[1])]
    return new_df.sort_values(by=['Room', 'Beg'])

def room_occupancy_on_day(_df, _room, _day):
    new_df = room_occupancy(_df, room=_room, day=_day)
    new_df = new_df.loc[:,['Beg','Title','Prof']]
    new_df = new_df.join(pandas.DataFrame(index=display_start_time),on='Beg', how='right').sort_values(by='Beg').reindex()
    new_df = new_df.fillna("").to_numpy()
    return [_day, _room], new_df

def pretty_print(df, _rooms, _days):
    num_col = 2 # adjust if more than instructor and class are needed
    header_array = numpy.full((2, len(_days)*len(_rooms)*(num_col + 1)), "", dtype=numpy.dtype('<U100'))
    display_array = None
    
    for i, day in enumerate(days):
        header_array[0,i*len(_rooms)*(num_col + 1)] = day
        for j, room in enumerate(_rooms):
            header_array[1,i*len(_rooms)*(num_col + 1)+j*(num_col + 1)] = ''.join(room)
            key, _array = room_occupancy_on_day(df, room, day)
            if display_array is None:
                display_array = _array
            else:
                # pad the smaller array so hstack can work
                pad_to = max(display_array.shape[0], _array.shape[0])
                display_array = numpy.pad(display_array, ((0, pad_to - display_array.shape[0]), (0,0)), constant_values='')
                _array = numpy.pad(_array, ((0, pad_to - _array.shape[0]), (0,0)), constant_values='')
                display_array = numpy.hstack((display_array, _array))

    return numpy.vstack((header_array, display_array))

if __name__ == '__main__':
    """ IMPORTANT
    whatever is downloaded from Section Tally MUST be resaved as a true .xls
    Section Tally outputs a broken xml file as far as I can tell
    """

    # parse section tally
    # merge overlapping course entries ?
    # reorganize into full printable structure
    # apply filters
    # drop unwanted entries
    config_path = 'engr_rooms_config.json'
    with open(config_path, 'r') as f:
        config_dict = json.load(f)

    room_list = config_dict['rooms']
    course_dict = config_dict['courses']
    faculty_list = config_dict['faculty']

    section_tally_target = 'section_tally_f23_resave.xls' 
    intermediate_output = 'section_tally_f23_parsed.xlsx'
    final_pretty_output = 'test_pretty_output.xlsx'
    days = ['M', 'T', 'W', 'R', 'F']

    df = parse_section_tally(section_tally_target)
    df = map_course_names(df, course_dict) # exact names only for now
    df = instructor_last_names(df)
    save_intermediate = True
    if save_intermediate:
        save_to_excel(df, intermediate_output)
    drop_names = False
    if drop_names:
        df = drop_names_not_in(df, faculty_list)
    pretty_array = pretty_print(df, room_list, days)
    save_to_excel(pandas.DataFrame(pretty_array), final_pretty_output)
