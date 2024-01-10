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
    
    room_col_re = r'^(?P<days>\S+)\s+(?P<Beg>\S+)\s+(?P<End>\S+)\s+(?P<Bldg>\S*)\s+(?P<Room>\S*)\s+(?P<Type>\S+)' 
    prof_lname_re = r'((?P<a>[^,]*)(?P<c1>,\s).*)?\s?((?P<b>[^,]*)(?P<c2>,\s).*)?\s?((?P<c>[^,]*),?.*)?'
    prof_lname_re_val = r'\g<a>\g<c1>\g<b>\g<c2>\g<c>'
    room_col = df.columns[-1] # former col I

    # filter out rows with empty room_col
    # split timeblock data
    df1 = df[room_col].str.split('\n', expand=True)\
        .stack()\
        .str.extractall(room_col_re)\
        .reset_index()\
        .rename(columns={"level_0": "index"})
    df1_mult = df1.loc[df1['days'].str.len() > 1]
    df1_single = df1.loc[df1['days'].str.len() <= 1]
    df_d = df1_mult["days"].str.extractall(r'(\S)').reset_index()
    # .reset_index(names=["orig", "l1", "l2", "l3"]).drop(axis = 1, columns = ["l1", "l2", "l3"])
    df1_mult = df1_mult.merge(df_d, left_index=True, right_on='level_0', how = 'inner')\
        .drop(axis = 1, columns = ["days", "level_0","level_1", "match_x", "match_y"])\
        .rename(columns={0: "Day"})
    df1_single = df1_single.rename(columns={"days": "Day"})\
        .drop(axis=1, columns=["level_1", "match"])
    df1 = pandas.concat([df1_mult, df1_single])
    df = df.drop(axis = 1, columns = "Day  Beg   End   Bldg Room  (Type)")\
        .merge(df1, left_index=True, right_on="index", how = 'inner')
    # filter to only Rowan Hall or Eng. Hall. This has the side effect of dropping Online courses or courses with unmarked rooms
    df = df[df['Bldg'].isin(['ROWAN', 'ENGR'])]\
        .drop(axis = 1, columns= 'index')\
        .reset_index(drop=True)
    df['Prof'] = df['Prof'].replace(regex=True,
                                    to_replace = prof_lname_re,
                                    value = prof_lname_re_val)\
        .replace(regex=True,
                 to_replace=r',\s?$',
                 value='')
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
    for t in display_start_time:
        if len(new_df.loc[new_df['Beg'] == t]) > 1: # squash overlapped sections
            temp = new_df.index[new_df['Beg'] == t].tolist()
            for i in temp[1:]:
                new_df.loc[temp[0],'Title'] += ' & ' + new_df.loc[i, 'Title']
                new_df.drop(index = i, inplace= True)
    new_df = new_df.fillna("").to_numpy()
    return [_day, _room], new_df

def pretty_print(df, _rooms, _days, _num_col):
    filter_df = df[['Title', 'Prof', 'Beg', "End", "Day"]]
    filter_df['BldgRoom'] = df['Bldg'] + '-' + df['Room']
    buildings_rooms = set([b + '-' + r for b,r in _rooms])
    filter_df = filter_df[filter_df.BldgRoom.isin(buildings_rooms)]
    def _pad_end(_array, final_axis):
        return numpy.pad(_array, ((0, final_axis - _array.shape[0]), (0,0)), constant_values='')
    
    num_col = _num_col
    header_array = numpy.full((2, len(_days)*len(_rooms)*(num_col)), "", dtype=numpy.dtype('<U100'))
    display_array = None

    for i, day in enumerate(display_days):
        header_array[0,i*len(_rooms)*(num_col)] = day
        for j, room in enumerate(_rooms):
            header_array[1,i*len(_rooms)*(num_col)+j*(num_col)] = ''.join(room)
            key, _array = room_occupancy_on_day(df, room, day)
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
    # start_times = numpy.pad(start_times, (1, 0), constant_values='')
    save_to_excel(pandas.DataFrame(display_array), 'test_output.xlsx')
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
    section_tally_target = 'section_tally_f24_resave.xls' 
    intermediate_output = 'section_tally_f24_parsed.xlsx'
    final_pretty_output = config_path.split('_')[0] + '_room_schedule_output.xlsx'
    display_days = ['M', 'T', 'W', 'R', 'F']
    num_col = 2 # adjust if more than instructor and class are needed

    save_intermediate_df = True
    drop_unknown_names = False

    df = parse_section_tally(section_tally_target)
    df = map_course_names(df, course_dict) # exact names only for now
    if save_intermediate_df:
        save_to_excel(df, intermediate_output)
    if drop_unknown_names:
        df = drop_names_not_in(df, faculty_list)
    pretty_array = pretty_print(df, room_list, display_days, num_col)
    # could leverage python xlsx writers to pretty up the code in script instead of this
    save_to_excel(pandas.DataFrame(pretty_array), final_pretty_output)
