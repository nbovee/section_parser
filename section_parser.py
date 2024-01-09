#!/usr/bin/python

import pandas
import numpy
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
        # df = pandas.read_xml(target) # doesn't play nice
    except Exception as e:
        print("Tip: Section Tally files must be resaved as true .xls first.")
        print(e)
        exit(1)
    
    room_col_re = r'^(?P<days>\S+)\s+(?P<start>\S+)\s+(?P<end>\S+)\s+(?P<building>\S*)\s+(?P<room>\S*)\s+(?P<type>\S+)' 
    room_col = df.columns[-1] # former col I

    # filter out rows with empty room_col
    # split timeblock data
    df1 = df[room_col].str.split('\n', expand=True)
    df1 = df1.stack()
    df1 = df1.str.extractall(room_col_re)
    df1 = df1.reset_index().rename(columns={"level_0": "index"})
    df1_mult = df1.loc[df1['days'].str.len() > 1]
    df1_single = df1.loc[df1['days'].str.len() <= 1]
    df_d = df1_mult["days"].str.extractall(r'(\S)').reset_index()
    # .reset_index(names=["orig", "l1", "l2", "l3"]).drop(axis = 1, columns = ["l1", "l2", "l3"])
    df1_mult = df1_mult.merge(df_d, left_index=True, right_on='level_0', how = 'inner')\
        .drop(axis = 1, columns = ["days", "level_0","level_1", "match_x", "match_y"]).rename(columns={0: "day"})
    df1_single = df1_single.rename(columns={"days": "day"}).drop(axis=1, columns=["level_1", "match"])
    df1 = pandas.concat([df1_mult, df1_single])
    df = df.drop(axis = 1, columns = "Day  Beg   End   Bldg Room  (Type)")
    df = df.merge(df1, left_index=True, right_on="index", how = 'inner')
    df1 = df1.rename(columns={0: room_col})
    df1 = df1.drop(axis=1, labels='level_1')
    # regex roomcol with below

    # split days data

    # reset index
    # split timeblock data (find \n in col I and make new entry)
    df1 = df[room_col].str.split('\n', expand=True)\
        .stack()\
        .reset_index(level=1)\
        .rename(columns={0: room_col})\
        .drop(axis=1, labels='level_1')
    # merge back in
    df = df.drop(axis=1, labels=room_col)\
        .merge(df1, left_index=True, right_index=True)\
        .reset_index(drop=True)
    # split room_col to individual col (split is strange on classes without rooms such as Online due to .split() functionality)
    df2 = df[room_col].str.split(expand=True)\
        .set_axis(room_col.split(), axis=1)
    # merge back in
    df = df.merge(df2, left_index=True, right_index=True)\
        .drop(axis=1, labels=room_col)

    # filter to only Rowan Hall or Eng. Hall. This has the side effect of dropping Online courses or courses with unmarked rooms
    df = df[df['Bldg'].isin(['ROWAN', 'ENGR'])]
    # handle multiple profs, '\n' to '; \n'
    df['Prof'] = df['Prof'].str.rstrip()#.replace('\n',';', regex=True)
    return df

def save_to_excel(dataframe, filename):
    with pandas.ExcelWriter(filename) as writer:
        dataframe.to_excel(writer, sheet_name='parsed', index=False, header = False)

def map_course_names(df, _dict):
    df['Title'] = df['Title'].str.strip() # catches some bad formatting from ST
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
    # new_df = _df.filter(['CRN','Prof'], axis=1)
    # new_df = new_df['Prof'].str.strip()
    # print(_df.to_markdown())
    lname_re = "^\s*([^\s,]+)"
    _tempppp = _df['Prof'].str.extractall(lname_re, re.MULTILINE)
    print(_tempppp.to_markdown())
    print(_tempppp.iat[188-27, 0])
    new_df = new_df.str.split('\n', expand = True).fillna('')
    for colname, coldata in new_df.items():
        temp = coldata.str.split(',', expand = True).drop(columns=[1])
        new_df[colname] = temp.fillna("")
    
    # if we find out how to use the string join properly, this can be vastly simplified
    new_df['Prof'] = new_df[0] + ', ' + new_df[1] #+ ', ' + new_df[2]
    new_df = new_df.drop(columns = [0, 1])#, 2])
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

def pretty_print(df, _rooms, _days, _num_col, overlaps):

    def _pad_end(_array, final_axis):
        return numpy.pad(_array, ((0, final_axis - _array.shape[0]), (0,0)), constant_values='')
    
    num_col = _num_col + overlaps
    header_array = numpy.full((2, len(_days)*len(_rooms)*(num_col)), "", dtype=numpy.dtype('<U100'))
    display_array = None

    for i, day in enumerate(days):
        header_array[0,i*len(_rooms)*(num_col)] = day
        for j, room in enumerate(_rooms):
            header_array[1,i*len(_rooms)*(num_col)+j*(num_col)] = ''.join(room)
            key, _array = room_occupancy_on_day(df, room, day)
            if not overlaps:
                start_times = _array[:,:1]
                _array = _array[:,1:]
            if display_array is None:
                display_array = _array
            else:
                # pad the smaller array so hstack can work
                pad_to = max(display_array.shape[0], _array.shape[0])
                display_array = _pad_end(display_array, pad_to)
                _array = _pad_end(_array, pad_to)
                display_array = numpy.hstack((display_array, _array))

    if not overlaps: # this will fail if there are overlaps
        # start_times = numpy.pad(start_times, (1, 0), constant_values='')
        print(start_times)
        print(display_array)
        display_array = numpy.hstack((start_times, display_array))
        header_array = numpy.pad(header_array, ((0, 0), (1, 0)), constant_values='')
    return numpy.vstack((header_array, display_array))

if __name__ == '__main__':
    """ IMPORTANT
    whatever is downloaded from Section Tally MUST be resaved as a true .xls
    Section Tally outputs a broken xml file as far as I can tell
    """
    # time fields are heavily duplicated since some courses over scheduled at the same time same room
    # only practical way to clean that up without a lot of programming is manual toggle
    config_path = 'exeed_config.json'
    with open(config_path, 'r') as f:
        config_dict = json.load(f)

    room_list = config_dict['rooms']
    course_dict = config_dict['courses']
    faculty_list = config_dict['faculty']
    course_overlaps = False
    section_tally_target = 'section_tally_f24_resave.xls' 
    intermediate_output = 'section_tally_f24_parsed.xlsx'
    final_pretty_output = config_path.split('_')[0] + '_room_schedule_output.xlsx'
    days = ['M', 'T', 'W', 'R', 'F']
    num_col = 2 # adjust if more than instructor and class are needed

    save_intermediate_df = False
    drop_unknown_names = False

    df = parse_section_tally(section_tally_target)
    print(df.to_markdown())
    df = map_course_names(df, course_dict) # exact names only for now
    df = instructor_last_names(df)
    
    if save_intermediate_df:
        save_to_excel(df, intermediate_output)
    if drop_unknown_names:
        df = drop_names_not_in(df, faculty_list)
    pretty_array = pretty_print(df, room_list, days, num_col, course_overlaps)
    # could leverage python xlsx writers to pretty up the code in script instead of this
    save_to_excel(pandas.DataFrame(pretty_array), final_pretty_output)
