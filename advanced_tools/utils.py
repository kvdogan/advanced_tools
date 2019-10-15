"""
Python tools and algorithms gathered through out the development projects and tutorials.

Sections:
1. File Read/Write/Convert/Save Operations
2. Pandas Utils
3. Path Operations for File/Folder/System
4. Algorithms for Hierarchical Structures
5. win32 COM utilities (not in use)
"""

import os
import subprocess
import json
import csv
import collections
import xlrd
import xlsxwriter
import pandas as pd
from itertools import chain
from six import string_types
from collections import defaultdict

# ###################################### Section 1 ##################################### #


def json_to_csv(input_file_path, output_file_path, sep=';'):
    """
    Python tool for converting nested json files to simple 2D csv file with given delimiter.
    :param:: input_file_path : Full path to input json file.
    :param:: output_file_path : Full path to output csv file.
    :param:: sep : Delimiter, default to ';'.
    """
    def json_to_dicts(json_str):
        try:
            objects = json.loads(json_str)
        except json.decoder.JSONDecodeError:
            objects = [json.loads(l) for l in json_str.split('\n') if l.strip()]

        return [dict(to_keyvalue_pairs(obj)) for obj in objects]

    def to_keyvalue_pairs(source, ancestors=[], key_delimeter='_'):
        def is_sequence(arg):
            return (not isinstance(arg, string_types)) and (hasattr(
                arg, "__getitem__") or hasattr(arg, "__iter__"))

        def is_dict(arg):
            return isinstance(arg, dict)

        if is_dict(source):
            result = [to_keyvalue_pairs(source[key], ancestors + [key]) for key in source.keys()]
            return list(chain.from_iterable(result))
        elif is_sequence(source):
            result = [
                to_keyvalue_pairs(item, ancestors + [str(index)])
                for (index, item) in enumerate(source)
            ]
            return list(chain.from_iterable(result))
        else:
            return [(key_delimeter.join(ancestors), source)]

    def dicts_to_csv(source, output_file, sep):
        def build_row(dict_obj, keys):
            return [dict_obj.get(k, "") for k in keys]

        keys = sorted(set(chain.from_iterable([o.keys() for o in source])))
        rows = [build_row(d, keys) for d in source]

        cw = csv.writer(output_file, delimiter=sep, lineterminator='\n')
        cw.writerow(keys)
        for row in rows:
            cw.writerow([c if isinstance(c, string_types) else c for c in row])

    with open(input_file_path) as input_file:
        json_file = input_file.read()
    dicts = json_to_dicts(json_file)
    with open(output_file_path, "w") as output_file:
        dicts_to_csv(dicts, output_file, sep)


def checkfile(path):
    """Check file in the path, and extend with '(1)' like numbering system if there exists.
    That helps to avoid overwriting and existing file.
    """
    path = os.path.expanduser(path)

    if not os.path.exists(path):
        return path

    root, ext = os.path.splitext(os.path.expanduser(path))
    dir = os.path.dirname(root)
    fname = os.path.basename(root)
    candidate = fname + ext
    index = 1
    ls = set(os.listdir(dir))
    while candidate in ls:
        candidate = "{}({}){}".format(fname, index, ext)
        index += 1
    return os.path.join(dir, candidate)


def read_csv_to_lol(full_path, sep=";"):
    """
    Read csv file into lists of list.
    Make sure to have a empty line at the bottom
    """
    with open(full_path, 'r') as ff:
        # read from CSV
        data = ff.readlines()
    # New line at the end of each line is removed
    data = [i.replace("\n", "") for i in data]
    # Creating lists of list
    data = [i.split(sep) for i in data]
    return data


def read_excel_to_lol(fname, sheet_index=0):
    """
    Read excel file into lists of list.
    Make sure to indicate sheet index if it is not the first sheet
    """
    wb = xlrd.open_workbook(fname)
    sh = wb.sheet_by_index(sheet_index)
    return [sh.row_values(i) for i in range(sh.nrows)]


def write_to_txt(list_item, full_path_txt):
    """
    Create a txt file for output to the report folder of the project.
    :param list_item
    :param full_path_txt
    :return:
    """
    if isinstance(list_item, list):
        with open(full_path_txt, 'w') as ff:
            for i in list_item:
                ff.write("{}\n".format(i))
    else:
        raise TypeError("Please use a python list for write into txt file")


def read_from_txt(file):
    """Read txt file as a list, each line becomes list item."""
    readings = []
    with open(file, 'r') as ff:
        for i in ff.readlines():
            readings.append(i.replace('\n', ""))
    return readings


def combine_txt_files(folder, sep='\t', encoding='latin1', skip_rows=0):
    r"""
    Combine txt files in a folder into one file, first column is written as file name.
    :param folder: Path to folder of files
    :param sep: Separator for reading, default is '\t'
    :param encoding: default is latin1
    :param skiprows: Skiprows of header repeating all over in each file
    :return:

    """
    files = get_filepaths(folder)
    with open('Combined_output.txt', 'a', encoding='latin1') as output:
        for filex in files:
            with open(filex, 'r', encoding='latin1') as source:
                lines = source.readlines()
                for line in lines[skip_rows:]:
                    output.writelines(source.name + '\t' + line)

# ###################################### Section 2 ##################################### #


def combine_multiple_csv_into_excel(full_path_to_folder=None, sep='\t', encoding='latin1'):
    r"""
    Combine csv files that can be converted to Dataframe and have same exact structure.
    :param full_path_to_folder:
    :param sep: Text separator, default is '\t'
    :param encoding: Text encoding, default is 'latin1'
    :return: excel file with one extra column showing the name of the file.
    """
    csv_files = sorted(get_filepaths(full_path_to_folder))
    folder_name = os.path.split(full_path_to_folder)[1]  # For folder location and folder name

    df_base = pd.read_csv(csv_files[0], sep=sep, encoding=encoding, low_memory=False)
    df_base['File_Name'] = os.path.splitext(os.path.split(csv_files[0])[1])[0]

    for i in csv_files[1:]:
        df_temp = pd.read_csv(i, sep=sep, encoding=encoding, low_memory=False)
        file_name = os.path.splitext(os.path.split(i)[1])[0]
        df_temp['File_Name'] = file_name

        df_base = df_base.append(df_temp)

    df_base.to_excel('{}\\{}.xlsx'.format(full_path_to_folder, folder_name))


def split_worksheets(file):
    """
    :param file: Excel file to be split by its worksheets.
    :return:
    """
    dfs_to_split = pd.read_excel(file, None, encoding='latin1')
    # 'None' used as worksheet kwarg thus it could be read as Dataframe dict.
    dfs_to_split = collections.OrderedDict(sorted(dfs_to_split.items()))
    for k, v in dfs_to_split.items():
        export_file_name = os.path.join(
            os.path.split(file)[0], "{}.xlsx".format(k)
        )
        writer = pd.ExcelWriter(export_file_name, engine='xlsxwriter')
        v.to_excel(excel_writer=writer, sheet_name=k, index=False)
        writer.save()
        writer.close()


def dataframe_diff(df1, df2):
    """
    Give difference between two pandas dataframe.
             Date   Fruit   Num   Color
    9  2013-11-25  Orange   8.6  Orange
    8  2013-11-25   Apple  22.1     Red
    """
    df = pd.concat([df1, df2])
    df = df.reset_index(drop=True)

    # group by
    df_gpby = df.groupby(list(df.columns))

    # get index of unique records
    idx = [x[0] for x in df_gpby.groups.values() if len(x) == 1]

    # filter

    return df.reindex(idx)


def dataframe_countif(df, col):
    """
    :param df1: Dataframe
    :param col: Column to count
    """
    new_col = col + "_Count"
    df1 = df.copy()
    df1[new_col] = df1.groupby(col)[col].transform('count')
    return df1

# ###################################### Section 3 #####################################


def export_file_names(file_type=None):
    """
    Walk through the folder structures and creates excel file with hyperlink.
    to every single file
    :param file_type:
    :return:
    """
    from tkinter import messagebox, filedialog, Tk
    import warnings  # xlsx writer warning is eliminated
    window = Tk()
    window.wm_withdraw()
    warnings.filterwarnings("ignore")  # xlsx writer hyperlink for 255+ char
    folder = filedialog.askdirectory(title='Please choose the folder to extract file names')
    filenames = get_filepaths(folder)
    if file_type is None or not isinstance(file_type, str):
        filenames = [i for i in filenames if "$" not in i]
    else:
        filenames = [i for i in filenames if i.endswith(file_type) and "$" not in i]

    # Add workbook and worksheet
    wkbk = os.path.join(folder, '____Link2Files____.xlsx')
    wb1 = xlsxwriter.Workbook(wkbk)
    ws1 = wb1.add_worksheet('FileNames')
    # Add a general format
    bold = wb1.add_format({'bold': 1})
    # write the first row as heading
    ws1.write(0, 0, "File Location", bold)
    ws1.write(0, 1, "File Name", bold)
    ws1.write(0, 2, 'Hyperlink', bold)

    row = 0
    while row < len(filenames):
        splited_file_name = os.path.split(filenames[row])
        ws1.write(row + 1, 0, splited_file_name[0])
        ws1.write(row + 1, 1, splited_file_name[1])
        ws1.write_url(row + 1, 2, filenames[row], string='Open File')  # Implicit format.
        row += 1

    wb1.close()

    messagebox.showinfo(title="Complete",
                        message="Done!",
                        detail="")


def get_folder_structure(directory=None, file_type=""):
    """
    Create a nested dictionary that represents the folder structure of rootdir.
    :param directory:
    :param file_type:
    :return:
    """
    alld = {'': {}}
    for dirpath, dirnames, filenames in os.walk(directory):
        d = alld
        dirpath = dirpath[len(directory):]
        for subd in dirpath.split(os.sep):
            based = d
            d = d[subd]
        if dirnames:
            for dn in dirnames:
                d[dn] = {}
        else:
            based[subd] = [i for i in filenames if i.endswith(file_type) and "$" not in i]
    return alld['']


def get_filepaths(rootdir=None, file_type='all', flat=True):
    r"""
    Advanced tool for getting file paths from nested folders.

    Arguments ::
    =================
    :param rootdir          : Fullpath to directory, string, default is None.
    :param file_type        : File extension to look up, string or list, use dot '.' in extension
    :param flat             : Return either flat list consist of fullpath of files from
                              nested folders if 'flat' is True.
    :return                 : list

    >>  In case of flat argument is True:
    >   ['file1_fullpath', 'file2_fullpath', ....]
    >   ['C:\\temp\\xxx.txt',r'C:\\temp\\temp2\\xxx.txt', ....]

    >>  In case of flat argument is False:
        Returns a list of tuples consist of filename, full path to folder and index in rootfir

    >   [(file1_name, file1_parentfolder_path, file1_peers_number)]
    >   [
            ("xxx.txt", r'C:\\temp', 2),
            ('xxx.txt',r'C:\\temp\\temp2',5)
        ]
    """
    if file_type == 'all' or file_type == '':
        file_paths = [
            (file, root) for root, directories, files in os.walk(rootdir)
            for file in files if "$" not in file
        ]
    else:
        file_paths = [
            (file, root) for root, directories, files in os.walk(rootdir)
            for file in files if os.path.splitext(file)[1] in file_type and "$" not in file
        ]

    if flat:
        return [os.path.join(root, file) for file, root in file_paths]
    else:
        return [i + (file_paths.index(i) + 1,) for i in file_paths]


def prepare_and_sortphotos(
    source_dir=r"./Pictures",
    destination_dir=r"./Pictures/ArrangedPictures",
    extension_to_fix='.HEIC',
    rename=True
):
    """This is a simple wrapper function for sortphotos library. Prepares images by
    converting .HEIC extensions to jpg if necessary.

    For more detail implementations check sortphotos library options.

    Keyword Arguments:
        source_dir {str} -- Directory with photos to arrange (default: {r"./Pictures"})
        destination_dir {str} -- Directory for arranged photos, in case of not existing,
        it creates automatically (default: {r"./Pictures/ArrangedPictures"})
        extension_to_fix {str} -- Extension to convert to .jpg (default: {".HEIC"})

    """

    pics = get_filepaths(source_dir)

    for i in pics:
        if os.path.splitext(i)[1] == extension_to_fix:
            os.rename(i, os.path.join(os.path.splitext(i)[0] + ".jpg"))
        else:
            pass
    if rename:
        subprocess.call(
            "sortphotos {} {} --rename=%Y_%m_%d_%H%M".format(source_dir, destination_dir)
        )
    else:
        subprocess.call("sortphotos {} {}".format(source_dir, destination_dir))

# ###################################### Section 4 ##################################### #


def hierarchy_tree(table, output_file=None, on_screen=False):
    """
    :param table:   table is a list of tuples of two, first item in tuple is child tag
                    while the second item is parent.
                    page_ids = [
                        (22, 4), (45, 1), (1, 1), (4, 4),
                        (566, 45), (7, 7), (783, 566), (66, 1), (300, 8),
                        (8, 4), (101, 7), (80, 22), (17, 17), (911, 66)
                    ]
    :param output_file: Absolute path for desired txt output_file
    :param on_screen: Default is False, if it is True, it prints whole hierarchy on screen
    """
    output = open(output_file, "w", encoding="utf-8")
    nodes, roots = defaultdict(set), set()

    for child, parent in table:
        if child == parent:
            roots.add(child)
        else:
            nodes[parent].add(child)

    # nodes now looks something like this:
    # {1: [45, 66], 66: [911], 4: [22, 8], 22: [80],
    #  7: [101], 8: [300], 45: [566], 566: [783]}

    def display(item, nodes, level):
        if on_screen is False and output_file is not None:
            output.writelines('%s%s%s' % ('\t|' * level, '\\_', item + '\n'))
            for child in sorted(nodes.get(item, [])):
                display(child, nodes, level + 1)
        else:
            print('%s%s%s' % ('\t' * level, '\\_', item))
            for child in sorted(nodes.get(item, [])):
                display(child, nodes, level + 1)

    for item in sorted(roots):
        display(item, nodes, 0)
    output.close()


def outlined_hierarchy(txtfile, sysname="HVAC_sample", sysno="97_sample",
                       wkbk="Outlined_Hierarchy_sample.xlsx", ws="Hierarchy"):
    """
    Create a hierarchical structure from the given file by looking parent and child relationship
    Arguments:
        txtfile {[type]} -- Structured txt file which is the output of hierarchy tree algorithm

    Keyword Arguments:
        sysname {str} -- Name of the system (default: {"HVAC_sample"})
        sysno {str} -- Number of the system (default: {"97_sample"})
        wkbk {str} -- Name for output excel file (default: {"Outlined_Hierarchy_sample.xlsx"})
        ws {str} -- Name for excel sheey (default: {"Hierarchy"})
    """
    ff = open(txtfile, "r", encoding="utf-8")
    rows = ff.readlines()
    ff.seek(0)
    ff.close()
    # Add workbook and worksheet
    wb = xlsxwriter.Workbook(wkbk)
    ws1 = wb.add_worksheet(ws)

    # Add a general format
    bold = wb.add_format({'bold': 1})
    level_text = wb.add_format({'bold': 1, 'bg_color': 'yellow'})

    # Freeze Top Pane
    ws1.freeze_panes(1, 0)

    # write the first row as heading (projects info and child tag levels)
    ws1.write(0, 0, "System: {}, System No:{}".format(sysname, sysno), bold)
    total_level = max(list(map(lambda x: x.count('|'), rows)))
    for i in range(2, total_level + 2):
        ws1.write(0, i - 1, "Level_{}".format(i), level_text)

    rowfor = 1
    colfor = 0
    while rowfor < len(rows):
        ws1.set_row(
            rowfor, None, None, {'level': colfor + rows[rowfor - 1].count("|"), 'hidden': False}
        )
        rowfor += 1

    row = 1
    col = 0
    while row < len(rows):
        ws1.write(row, col + rows[row - 1].count("|"), rows[row - 1])
        row += 1
    wb.close()


def get_hierarchy_as_list(
    main_df, tag_list,
    lookup_for='children',
    exclude_list=None,
    sub_level=None,
    tag_column='Functional Location',
    parent_column='Superior functional location',
    print_details=True
):
    r"""
    Retrieve parent tags or children tags as a dataframe for a given tag lists

    Arguments:
        main_df {pandas.Dataframe} -- Tag
        tag_list {list} -- Reference Tag list

    Keyword Arguments:
        lookup_for {str} -- Either 'children' or 'parents' at a time (default: {'parents'})
        exclude_list {list} -- List of tags that will be excluded from final output
        (default: {None})
        sub_level {int} -- Level of parent or children tags (default: {None})
        tag_column {str} -- Name of the tag column in main_df (default: {'Functional Location'})
        parent_column {str} -- Name of the parent tag column in main_df
        (default: {'Superior functional location'})

    Returns:
        [pandas.Dataframe] -- [Output of desired list of parents or children]
    """
    if lookup_for == 'children':
        tag_field = tag_column
        parent_field = parent_column
        sublevel_direction = lambda x: -abs(x)
    elif lookup_for == 'parents':
        tag_field = parent_column
        parent_field = tag_column
        sublevel_direction = lambda x: abs(x)
    else:
        raise (AttributeError("lookup_for keyword argument must be either 'parents' or 'children'"))

    # Extracting initial dataframe as sublevel=1 from df['sap'] with initial taglist.
    df_out = main_df[main_df[tag_column].isin(tag_list)].copy()
    df_out['Sublevel'] = 0
    tags = df_out[tag_field].tolist()

    if sub_level is not None:
        ctr = 1
        if print_details:
            print("Total number of tags: " + str(len(tags)) + " at level-" + str(ctr - 1))
        # Extracting all upto given sublevel starting from the initial taglist
        while ctr <= sub_level:
            df_temp = main_df[main_df[parent_field].isin(tags)]
            df_temp['Sublevel'] = sublevel_direction(ctr)
            df_out = df_out.append(df_temp)
            tags = df_temp[tag_field].tolist()
            if print_details:
                print("Total number of tags: " + str(len(tags)) + " at level-" + str(ctr))
            ctr += 1
    else:
        ctr = 1
        if print_details:
            print("Total number of tags: " + str(len(tags)) + " at level-" + str(ctr - 1))
        # Extracting all the sublevel starting from the initial taglist
        while len(tags) > 0:
            df_temp = main_df[main_df[parent_field].isin(tags)]
            df_temp['Sublevel'] = sublevel_direction(ctr)
            df_out = df_out.append(df_temp)
            tags = df_temp[tag_field].tolist()
            if print_details:
                print("Total number of tags: " + str(len(tags)) + " at level-" + str(ctr))
            ctr += 1

    # Drop duplicates with the subset of Functional Location
    df_out = df_out.drop_duplicates(subset=[tag_column], keep='last')

    # Filtering out tags which already exist in AlignIT.
    if exclude_list:
        df_out = df_out[~df_out[tag_column].isin(exclude_list)]

    return df_out


def apex_workorders(
    hierarchyDF=None,
    lookupDF=None,
    taglist=None,
    tag_column='Functional Location',
    parent_column='Superior functional location',
    parents=dict(include=True, level=None, include_children_of_parents=False, cop_level=None),
    children=dict(include=True, level=None),
    sisters=dict(include=True, include_children_of_sisters=False, cos_level=None),
    export=False,
    groupby_columns=['Functional Location', 'Order', 'Description'],
):
    """
    Compiles workorder based on desired hierarchical way, with the aid of get_hierarchy_as_list
    utility function from advanced_tools and pandas library.

    Keyword Arguments:
        hierarchyDF {pandas.DataFrame} -- DataFrame that represents hierarchy with columns of
            tag_column and parent columns (default: {None})
        lookupDF {pandas.DataFrame} -- Lookup DataFrame matching tag column in hierarchyDF
            (default: {None})
        taglist {list} -- List of tags to merge from lookupDF (default: {None})
        tag_column {str} -- Name of the tag column (default: {'Functional Location'})
        parent_column {str} -- Name of the parent column
            (default: {'Superior functional location'})
        parents {dict} -- Dict of arguments for get_hierarchy_as_list function (default:
            {dict(include=True, level=None, include_children_of_parents=False, cop_level=None)})
        children {dict} -- Dict of arguments for get_hierarchy_as_list function
            (default: {dict(include=True, level=None)})
        sisters {dict} -- Dict of arguments for get_hierarchy_as_list function
            (default: {dict(include_children_of_sisters=False, cos_level=None)})
        export {bool} -- Export groupby output to excel (default: {False})
        groupby_columns {list} -- Name of the columns to groupby excel output
            (default: {['Functional Location', 'Order', 'Description']})

    Returns:
        {pandas.DataFrame} -- DataFrame object of compiled Workorders
    """
    if parents['include']:
        ptags = get_hierarchy_as_list(
            main_df=hierarchyDF, tag_list=taglist,
            tag_column=tag_column, parent_column=parent_column,
            lookup_for='parents', sub_level=parents['level'], print_details=False
        )
        if parents['include_children_of_parents']:
            coptags = get_hierarchy_as_list(
                main_df=hierarchyDF, tag_list=ptags[tag_column].tolist(),
                tag_column=tag_column, parent_column=parent_column,
                lookup_for='children', sub_level=parents['cop_level']
            )

            ptags = pd.concat([ptags, coptags])
    else:
        ptags = pd.DataFrame()

    if children['include']:
        ctags = get_hierarchy_as_list(
            main_df=hierarchyDF, tag_list=taglist,
            tag_column=tag_column, parent_column=parent_column,
            lookup_for='children', sub_level=children['level'],
            print_details=False
        )
    else:
        ctags = pd.DataFrame()

    if sisters['include']:
        first_parents = hierarchyDF.loc[hierarchyDF[tag_column].
                                        isin(taglist), parent_column].drop_duplicates().tolist()

        sister_tags = hierarchyDF.loc[hierarchyDF[parent_column].
                                      isin(first_parents), tag_column].unique().tolist()

        if not sisters['include_children_of_sisters']:
            stags = hierarchyDF[hierarchyDF[tag_column].isin(sister_tags)]
            stags['Sublevel'] = 0
        else:
            stags = get_hierarchy_as_list(
                main_df=hierarchyDF, tag_list=sister_tags,
                tag_column=tag_column, parent_column=parent_column,
                lookup_for='children', sub_level=sisters['cos_level'],
                print_details=False
            )
    else:
        stags = pd.DataFrame()

    ttags = pd.concat([ptags, stags, ctags])

    if len(ttags) > 0:
        ttags.drop_duplicates(inplace=True, keep='first')
    else:
        ttags = hierarchyDF[hierarchyDF[tag_column].isin(taglist)]
        ttags['Sublevel'] = 0

    wos = lookupDF[lookupDF[tag_column].isin(ttags[tag_column].tolist())]

    tos = wos.merge(
        ttags[['Functional Location', 'Location', 'Sublevel']], on="Functional Location", how='left'
    )

    if export:
        wos.fillna('NoValueAvailable', inplace=True)
        wos.groupby(groupby_columns).agg('count').to_excel(
            'WOs_for_{}_with_P({}-{})-CoP({}-{})-C({}-{})-S({}-CoS{}-{}).xlsx'.format(
                taglist[0],
                str(parents['include']),
                str(parents['level']),
                str(parents['include_children_of_parents']),
                str(parents['cop_level']),
                str(children['include']),
                str(children['level']),
                str(sisters['include']),
                str(parents['include_children_of_sisters']),
                str(parents['cos_level']),
            )
        )
    else:
        pass
    return tos
