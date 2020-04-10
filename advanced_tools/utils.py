"""
Python tools and algorithms gathered through out the development projects and tutorials.

Sections:
1. File Read/Write/Convert/Save Operations
2. Pandas Utils
3. Path Operations for File/Folder/System
4. Algorithms for Hierarchical Structures
5. Utility functions for xlrd library and read_spir function
"""

import collections
import csv
import json
import os
import re
import subprocess
import warnings  # xlsx writer warning is eliminated
from collections import defaultdict
from itertools import chain
from tkinter import Tk, filedialog, messagebox

import pandas as pd
import xlrd as xl
import xlsxwriter
from six import string_types

# ########################## File Read/Write/Convert/Save Operations ########################## #


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
            return (not isinstance(arg, string_types)
                   ) and (hasattr(arg, "__getitem__") or hasattr(arg, "__iter__"))

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


def write_lol_to_csv(output_csv=None, headers=None, data=None, seperator=";"):
    if not isinstance(data, list) or not isinstance(headers, list):
        raise (TypeError("Data must be lists of list and header must be plain list"))
    if not os.path.isfile(output_csv):
        with open(output_csv, 'w', newline='', encoding="utf-8") as output:
            csv_output = csv.writer(output, delimiter=seperator)
            csv_output.writerow(headers)
    with open(output_csv, 'a', newline='', encoding="utf-8") as output:
        csv_output = csv.writer(output, delimiter=seperator)
        csv_output.writerows(data)


def read_excel_to_lol(fname, sheet_index=0):
    """
    Read excel file into lists of list.
    Make sure to indicate sheet index if it is not the first sheet
    """
    wb = xl.open_workbook(fname)
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
        raise (TypeError("Please use a python list for write into txt file"))


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


# ####################################### Pandas Utils ######################################## #


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
        export_file_name = os.path.join(os.path.split(file)[0], "{}.xlsx".format(k))
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


def split_and_export_dataframe(df, nrows, sortby=None, output_name=None, export_csv=True):
    """SPlit dataframe with given row numbers and export it wither in csv or excel
    
    Arguments:
        df {pd.DataFrame} -- Pandas dataframe
        nrows {int} -- Number of rows to split the database
    
    Keyword Arguments:
        sortby {string} -- Sort DataFrame before splitting (default: {None})
        output_name {string} -- Output file name to save, input if not given (default: {None})
        export_csv {bool} -- Output file format, (default: {True}) ';' seperated csv file'
    """
    if sortby is not None:
        df.sort_values(by=sortby, inplace=True)
    n_iter = int(pd.np.ceil(len(df) / nrows))
    if output_name is None:
        output_name = input("Please name the output file to save: ")

    if export_csv:
        for i in range(n_iter):
            df[nrows * i:nrows * i +
               nrows].to_csv(checkfile(r'./{}.csv'.format(output_name)), sep=';', index=False)
    else:
        for i in range(n_iter):
            df[nrows * i:nrows * i +
               nrows].to_excel(checkfile(r'./{}.xlsx'.format(output_name)), index=False)


def convert_to_hyperlink(x):
    """
    Converts pandas column given in apply function to hyperlink for excel export

    Usage:
        df['Link'] = df['Link'].apply(convert_to_hyperlink)
    
    Returns:
        DataFrame -- Creates or modifies given DataFrame column
    """
    if "#" in x:
        return "'#' in file path is not allowed in Excel Hyperlinks"
    else:
        return f'=HYPERLINK("{x}", "Click to Open")'


# ########################## Path Operations for File/Folder/System ########################## #


def export_file_names(file_type=None, use_relative_path=True):
    """
    Walk through the folder structures and creates excel file with hyperlink.
    to every single file

    Keyword Arguments:
        file_type {string} -- File extension i.e. ".xls" (default: {None})
        use_relative_path {bool} -- Relative paths to selected folder (default: {True})
    """
    window = Tk()
    window.wm_withdraw()
    warnings.filterwarnings("ignore")  # xlsx writer hyperlink for 255+ char
    folder = filedialog.askdirectory(title='Please choose the folder to extract file names')
    filenames = get_filepaths(folder)
    if file_type is None or not isinstance(file_type, str):
        filenames = [i for i in filenames if "$" not in i]
    else:
        filenames = [i for i in filenames if i.lower().endswith(file_type) and "$" not in i]

    df = pd.DataFrame(data=filenames, columns=['Path'])
    df['FileName'] = df['Path'].apply(lambda x: os.path.split(x)[1])

    if use_relative_path:
        df['Link'] = df['Path'].apply(lambda x: os.path.relpath(x, folder))
    else:
        df['Link'] = df['Path']

    df['Link'] = df['Link'].apply(convert_to_hyperlink)

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(
        os.path.join(os.path.join(folder, '____Link2Files____.xlsx')),
        engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='Link2Files', index=False)
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['Link2Files']
    # Add some cell formats.
    custom_hyperlink_format = workbook.add_format({
        'font_color': 'blue',
        # 'bold':       1,
        'underline':  1,
        'font_size':  12,
    })
    # Note: It isn't possible to format any cells that already have a format such
    # as the index or headers or any cells that contain dates or datetimes.
    # Set the format but not the column width.
    worksheet.set_column('C:C', None, custom_hyperlink_format)
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    messagebox.showinfo(title="Complete", message="Done!", detail="")


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
            based[subd] = [i for i in filenames if i.lower().endswith(file_type) and "$" not in i]
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
            for file in files if os.path.splitext(file)[1].lower() in file_type and "$" not in file
        ]

    if flat:
        return [os.path.join(root, file) for file, root in file_paths]
    else:
        return [i + (file_paths.index(i) + 1, ) for i in file_paths]


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


# ########################## Algorithms for Hierarchical Structures ########################## #


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


def outlined_hierarchy(
    txtfile,
    sysname="HVAC_sample",
    sysno="97_sample",
    wkbk="Outlined_Hierarchy_sample.xlsx",
    ws="Hierarchy"
):
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
            rowfor, None, None, {
                'level': colfor + rows[rowfor - 1].count("|"),
                'hidden': False
            }
        )
        rowfor += 1

    row = 1
    col = 0
    while row < len(rows):
        ws1.write(row, col + rows[row - 1].count("|"), rows[row - 1])
        row += 1
    wb.close()


def get_hierarchy_as_list(
    main_df,
    tag_list,
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
            main_df=hierarchyDF,
            tag_list=taglist,
            tag_column=tag_column,
            parent_column=parent_column,
            lookup_for='parents',
            sub_level=parents['level'],
            print_details=False
        )
        if parents['include_children_of_parents']:
            coptags = get_hierarchy_as_list(
                main_df=hierarchyDF,
                tag_list=ptags[tag_column].tolist(),
                tag_column=tag_column,
                parent_column=parent_column,
                lookup_for='children',
                sub_level=parents['cop_level'],
                print_details=False
            )

            ptags = pd.concat([ptags, coptags])
    else:
        ptags = pd.DataFrame()

    if children['include']:
        ctags = get_hierarchy_as_list(
            main_df=hierarchyDF,
            tag_list=taglist,
            tag_column=tag_column,
            parent_column=parent_column,
            lookup_for='children',
            sub_level=children['level'],
            print_details=False
        )
    else:
        ctags = pd.DataFrame()

    if sisters['include']:
        first_parents = hierarchyDF.loc[hierarchyDF[tag_column].isin(taglist),
                                        parent_column].drop_duplicates().tolist()

        sister_tags = hierarchyDF.loc[hierarchyDF[parent_column].isin(first_parents),
                                      tag_column].unique().tolist()

        if not sisters['include_children_of_sisters']:
            stags = hierarchyDF[hierarchyDF[tag_column].isin(sister_tags)]
            stags['Sublevel'] = 0
        else:
            stags = get_hierarchy_as_list(
                main_df=hierarchyDF,
                tag_list=sister_tags,
                tag_column=tag_column,
                parent_column=parent_column,
                lookup_for='children',
                sub_level=sisters['cos_level'],
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
        ttags[['Functional Location', 'Location', 'Sublevel']],
        on="Functional Location",
        how='left'
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


# ######################### Utility functions for Python xlrd library ######################### #

def _get_cell_range(sheet_obj, start_row, start_col, end_row, end_col):
    """
    Get cell range in xlrd module as two level nested list.
    
    Arguments:
        sheet_obj {xlrd worksheet object} -- xlrd worksheet instance
        start_row {int} -- Number of start row
        start_col {int} -- Number of start column
        end_row {int} -- Number of last row
        end_col {int} -- Number of last column
    
    Returns:
        list -- Cell range as two level nested list
    """
    return [
        sheet_obj.row_slice(row, start_colx=start_col, end_colx=end_col + 1)
        for row in range(start_row, end_row + 1)
    ]


def _convert_empty_cells(sheet_obj, convert_to=None):
    """
    Create a list with all row numbers that contain data and loop through it.
    Create a list with all column numbers that contain data and loop through i
    
    Arguments:
        sheet_obj {xlrd worksheet object}
    
    Keyword Arguments:
        convert_to {str, int, None} -- (default: {None})
    """
    for r in range(0, sheet_obj.nrows):
        for c in range(0, sheet_obj.ncols):
            if sheet_obj.cell_type(r, c) == xl.XL_CELL_EMPTY:
                sheet_obj._cell_values[r][c] = convert_to


def _get_sheet_dimension(sheet_obj):
    """Return sheet dimension for quality validation issues
    
    Arguments:
        sheet_obj {xlrd worksheet object}
    
    Returns:
        dict -- Dictionary of 'spirname', 'maxcol', 'maxrow'
    """
    return {'spirname': sheet_obj.name, 'maxcol': sheet_obj.ncols, 'maxrow': sheet_obj.nrows}
    # print(f"{spir_sheet.name}\tmaxCol: {spir_sheet.ncols}\tmaxRow: {spir_sheet.nrows}")


def _find_pattern(sheet_obj, pattern):
    """
    Finds given regex pattern
    
    Arguments:
        sheet_obj {xlrd worksheet object}
        pattern {string} -- Regex pattern
    
    Returns:
        list -- List of (row, col) tuples that match regex pattern
    """

    row_col_list = []
    for r in range(0, sheet_obj.nrows):
        for c in range(0, sheet_obj.ncols):
            if re.search(pattern, str(sheet_obj.cell_value(r, c))):
                row_col_list.append(sheet_obj.cell_value(r, c))

    return row_col_list

    # if len(set(row_col_list)) == 1:
    #     return row_col_list[0]
    # elif len(set(row_col_list)) > 1:
    #     return sheet_obj.cell_value(*cell_for_spir)
    # else:
    #     return "NO_SPIR_REF_FOUND"


def _get_horizontal_range(sheet_obj, row=None, start_col=None):
    """
    Find last filled cell in row, followed by two empty cell at least.
    start_col==1 because requires validation of "Tag No." text
    
    Arguments:
        sheet_obj {xlrd worksheet object}
    
    Keyword Arguments:
        row {int} -- First Tag cell row (default: {4})
        start_col {int} -- First Tag cell column (default: {1})
    
    Returns:
        dict -- Dictionary of 'tag' and 'last_cell'
    """
    if "tag" in sheet_obj.cell_value(row, start_col-1).lower():
        ctr = 0
        cell_list = sheet_obj.row_slice(row, start_col)
        z = []
        while ctr < len(cell_list):
            if cell_list[ctr].value is not None:
                z.append(cell_list[ctr])
                ctr += 1
            elif cell_list[ctr].value is None and cell_list[ctr + 1].value is None:
                break
            else:
                z.append(cell_list[ctr])
                ctr += 1
        return {'tags': z[1:], 'last_cell': (row, start_col + ctr - 1)}
    else:
        raise (ValueError("Check First Tag Cell in the SPIR"))


# ########################## SPIR Codification Program and Utilities ########################## #


def codify_spir(
    path_to_spir='ENI_SPIR_Ref_form.xlsx', raise_error=True, tag_cell=(4, 2), mm_cell=(10, 29),
):
    """
    Read SPIR forms for ENI Goliat Project or another companies.
    First tag cell is manual input while last_tag_cell, first/last MM cell are found based
    on script. Scripts are stored in utils.py
    
    Keyword Arguments:
    path_to_spir {str}  -- Full path to SPIR file (default: {'ENI_SPIR_Ref_form.xlsx'})
    tag_cell {tuple}    -- First Tag Cell (default: {(4, 2)})
    mm_cell {tuple}     -- First Material Cell, this is used for overwriting if find_mm_column 
                            function is not returning right value (default: {(10, 20)})
    raise_error {bool}  -- Raised or printed errors getting silenced for multiple SPIRs,
                            error handling can be done in wrapper structure (default: {False})
    """
    wb = xl.open_workbook(path_to_spir)
    print(f"Codification started for: {os.path.split(path_to_spir)[1]}")
    # From this point and on it is Platform Dependent ------------>
    try:
        spir_sheet_name = [
            i for i in wb.sheet_names() if re.match(r"^spir$", i.strip(), re.IGNORECASE)][0]
        spir_sheet = wb.sheet_by_name(spir_sheet_name)
        _convert_empty_cells(spir_sheet)
    except IndexError:
        raise (NameError("There is no SPIR spreadsheet for found in the excel file"))

    try:
        cover_page = [i for i in wb.sheet_names() if re.match("front|cover", i, re.IGNORECASE)][0]
        coversheet = wb.sheet_by_name(cover_page)
        _convert_empty_cells(coversheet)
        spir_name = list(set(_find_pattern(coversheet, pattern=r".+-MC-.+")))[0]
    except IndexError as err:
        spir_name = os.path.split(path_to_spir)[1]
        if raise_error:
            print(f"{os.path.split(path_to_spir)[1]} : {err}")


    # Set reference cells, xlrd uses zero index row and column reference
    # xlrd works like range(x,y) function for col_values so increase last value by 1.
    ftc = tag_cell  # First tag cell coordinate
    # ltc = (4, 4)               # Last tag cell coordinate, kept for overwriting purposes
    # fmc = (10, 26)             # First material cell coordinate, kept for overwriting purposes
    # lmc = (42, 26)             # Last material cell coordinate, kept for overwriting purposes

    # Calculate number of spare parts in the SPIR form
    fmc = _find_mm_column(spir_sheet, first_mm=mm_cell)

    number_of_spares = len(spir_sheet.col_values(colx=fmc[1], start_rowx=fmc[0]))
    ltc = _get_horizontal_range(spir_sheet, row=ftc[0], start_col=ftc[1])['last_cell']
    lmc = (fmc[0] + number_of_spares - 1, fmc[1])  # Last material cell coordinate
    fqc = (fmc[0], ftc[1])  # First quantity cell
    lqc = (lmc[0], ltc[1])  # Last quantity cell

    # From this point and on it is Platform Independent ------------>

    # Read tag numbers as a simple list, row values works like range function so +1 on column
    tags = spir_sheet.row_values(rowx=ftc[0], start_colx=ftc[1], end_colx=ltc[1] + 1)

    # Create Tag-Spare quantity matrix table ($C$7:~)
    # Return two level nested list i.e. (row > columns)
    qty_tbl_rng = _get_cell_range(spir_sheet, fqc[0], fqc[1], lqc[0], lqc[1])
    qty_tbl = list(map(lambda x: list(map(lambda y: y.value, x)), qty_tbl_rng))

    # Key Algorithm 1
    # Create tag_table matrix using tag-spare quantity range (qty_tbl) ("C7:~")
    # Return three level nested list i.e. (table-wrapper > row > columns)
    tag_tbl = []
    ctr_row1 = 0
    while ctr_row1 < number_of_spares:
        ctr_col1 = 0
        temp_col_list = []
        while ctr_col1 < len(tags):
            if qty_tbl[ctr_row1][ctr_col1] is not None:
                temp_col_list.append(tags[ctr_col1])
            else:
                temp_col_list.append(None)

            ctr_col1 += 1
        tag_tbl.append(temp_col_list)
        ctr_row1 += 1

    # Filter None from tables, None.__ne__ to keep other falsify values such as 0, [], {} etc.
    tag_tbl = list(map(lambda x: list(filter(None.__ne__, x)), tag_tbl))
    qty_tbl = list(map(lambda x: list(filter(None.__ne__, x)), qty_tbl))

    # Create material number list (simple list)
    mat_tbl_rng = _get_cell_range(spir_sheet, fmc[0], fmc[1], lmc[0], lmc[1])
    mat_tbl_rng_value = [cell.value for row in mat_tbl_rng for cell in row]
    mat_tbl_rng_value = [i.strip().strip("'") if i is not None else i for i in mat_tbl_rng_value]

    # Replace trailng quote at the end of MM number
    pattern_mm = re.compile(r"[0-9]{15,20}", re.UNICODE)
    try:
        # First fill na values in mat_tbl with 999999... and then regex match the material
        mat_tbl = [i if i is not None else "99999999999999999999" for i in mat_tbl_rng_value]
        mat_tbl = list(map(lambda x: re.search(pattern_mm, x).group(0), mat_tbl))
    except (TypeError, AttributeError) as err:
        mat_tbl = [i if i is not None else "99999999999999999999" for i in mat_tbl_rng_value]
        if raise_error:
            print("Error while looking for material number regex match: ", err)
            # print("Some material number has wrong syntax, needs to be checked")

    # Validate lenght of tag, qty and material lists
    if len(tag_tbl) == len(mat_tbl) == len(qty_tbl):
        max_row_ctr = len(tag_tbl)
    else:
        # Python 3.6 new feature 'f string' is used
        raise (IndexError(
            f"""
            Inconsistent table!
            len(tag_tbl)==len(mat_tbl)==len(qty_tbl) condition is not confirmed
            Length of Tag table: {len(tag_tbl)}
            Length of Qty table: {len(qty_tbl)}
            Length of Mat table: {len(mat_tbl)}
            """
        ))

    # Key Algorithm 2
    # Replace any char other than hyphen around tag.
    # Split tag numbers written in same cell using ':;\n' separator.
    pattern = re.compile(r'[a-zA-Z0-9-]+', re.UNICODE)
    tag_tbl = list(map(lambda x: list(map(lambda y: re.findall(pattern, y), x)), tag_tbl))

    # Key Algorithm 3
    # Zip Tag number with material number and specified quantity as list of tuples of 3
    zipped_data = []
    ctr_row2 = 0
    while ctr_row2 < max_row_ctr:
        ctr_col2 = 0
        for i in tag_tbl[ctr_row2]:
            if len(i) == 1:
                zipped_data.append(
                    (i[0], qty_tbl[ctr_row2][ctr_col2], mat_tbl[ctr_row2], spir_name)
                )
                ctr_col2 += 1
            else:
                for j in i:
                    zipped_data.append(
                        (j, qty_tbl[ctr_row2][ctr_col2], mat_tbl[ctr_row2], spir_name)
                    )
                ctr_col2 += 1
        ctr_row2 += 1


    output_folder = os.path.join(os.path.split(path_to_spir)[0], "CodifiedFilesResults")

    if os.path.isdir(output_folder):
        pass
    else:
        os.makedirs(output_folder)

    tag_mat_qty_output = os.path.join(output_folder, "Tag_Mat_Qty.csv")
    spare_detail_output = os.path.join(output_folder, "Sparepart_Details.csv")

    write_lol_to_csv(
        output_csv=tag_mat_qty_output,
        headers=['Tag', 'Quantity', 'EniMM', 'SpirNo'], data=zipped_data, seperator=";")

    # Read from Spare part unit type to last column as a dataframe with util function
    spir_detail_export = _create_mm_table(
        sheet_obj=spir_sheet, srow=fmc[0], scol=fmc[1] - 19, erow=lmc[0], ecol=lmc[1] + 6
    )
    spir_detail_export['SpirNo'] = spir_name

    write_lol_to_csv(
        output_csv=spare_detail_output,
        headers=spir_detail_export.columns.tolist(),
        data=spir_detail_export.values.tolist(), seperator=";")

    os.rename(path_to_spir, os.path.join(output_folder, os.path.split(path_to_spir)[1]))


def codify_multiple_spir(tag_cell=(4, 2), mm_cell=(10, 29)):
    window = Tk()
    window.wm_withdraw()
    folder = filedialog.askdirectory(title='Please choose SPIR folder to codify')
    # filenames is obtained with os.scandir, because subfolder contains output files.
    fnames = [
        i.path for i in os.scandir(folder)
        if os.path.splitext(i.path)[1].lower() in ['.xls', '.xlsx', '.xlsm']
    ]

    if os.path.isfile(os.path.join(folder, '__Quality Report for Updated SPIR(s)__.xlsx')):
        pass
    else:
        # quality_assurance_check(folder)
        messagebox.showinfo(
            title="SPIR Quality Assurance",
            message="Consider checking SPIR qualities with the aid of quality_assurance_check()",
            detail=""
        )

    for i in fnames:
        try:
            codify_spir(path_to_spir=i, tag_cell=tag_cell, mm_cell=mm_cell, raise_error=False)
        except Exception as err:
            spir_errors = os.path.join(folder, 'SPIR_ERRORs.csv')
            with open(spir_errors, "a", encoding='utf-8') as report:
                report.write(os.path.split(i)[1]+";"+str(err) + "\n")
            continue

    messagebox.showinfo(
        title="Complete",
        message="Done! For possible errors check 'Quality Report' and 'Unstructured_SPIRs.txt'",
        detail=""
    )


def _find_mm_column(sheet_obj, pattern=r"^[0-9]{15,20}", first_mm=(None, None)):
    """
    Find MM columns with the help of regex pattern
    
    Arguments:
        sheet_obj {xlrd worksheet object}
    
    Keyword Arguments:
        pattern {regexp} -- (default: {r"^[0-9]{15,20}"})
        first_mm {tuple} -- Fallback value for first material number cell in case of unsuccessful
                            parsing (default: {None})
    
    Returns:
        tuple -- Tuple of cell cordinates
    """
    row_col_list = []
    for r in range(0, sheet_obj.nrows):
        for c in range(0, sheet_obj.ncols):
            if re.search(pattern, str(sheet_obj.cell_value(r, c))):
                row_col_list.append((r, c))
    seen = set()
    dups = set()
    for r, c in row_col_list:
        if c in seen:
            dups.add(c)
        seen.add(c)
    try:
        column = max(dups)
        row = min([r for r, c in row_col_list if c == column])
        return (row, column)
    except (TypeError, ValueError):
        print("Issue: MM number can't be fetched by find_mm_column method")
        return first_mm
    finally:
        pass


def _create_mm_table(sheet_obj, srow, scol, erow, ecol):
    """Get MM table, by putting mm number as index.
    
    Arguments:
        sheet_obj {xlrd worksheet object}
        srow {int} -- Start row for file
        scol {int} -- Start column for file
        erow {int} -- End row for file
        ecol {int} -- End column for file
    
    Returns:
        pandas Dataframe
    """
    table_columns = [
        "SpareUnitCode", "SparePartDescription", "LongText", "DetailDocumentNo",
        "DetailDocumentItemRef", "Material", "SupplierPartNo", "ManufacturerPartNo",
        "ManufacturerName", "SupplierRecommCommQty", "EngineeringRecommCommQty", "OrderedCommQty",
        "SupplierRecommOperationalQty", "EngineeringRecommOperationalQty", "OrderedOperationalQty",
        "SupplierRecommCapitalQty", "EngineeringRecommCapitalnQty", "OrderedCapitalQty",
        "MeasureUnit", "EniMM", "AchillesCode", "BatchManagement", "SerialNo", "UnitPriceNOK",
        "OperatinalSpareDeliveryTimeInMonth", "PreviouslyDelivered"
    ]

    # Read from Spare part unit type to last column
    range_for_df = _get_cell_range(
        sheet_obj, start_row=srow, start_col=scol, end_row=erow, end_col=ecol
    )
    # Convert range to xlrd values
    range_for_df = list(map(lambda x: list(map(lambda y: y.value, x)), range_for_df))

    # Read as Dataframe
    df = pd.DataFrame(range_for_df, columns=table_columns)

    # Pandas method for extracting regex match from column
    df['EniMM'] = df['EniMM'].str.extract("([0-9]{15,20})", expand=False)
    df['EniMM'] = df['EniMM'].fillna('99999999999999999999')
    return df


def quality_assurance_check(path_to_folder=r"./ENI_test_spirs", use_relative_path=True):
    """Use Utility functions to validate quality of SPIRs.
    
    Keyword Arguments:
        path_to_folder {path} -- Path to SPIR folder (default: {r"./ENI_test_spirs"})
    """
    fnames = get_filepaths(path_to_folder, file_type=['.xls', '.xlsx', '.xlsm'])
    fnames = [i for i in fnames if "Original" not in i]
    report_list = []
    for file in fnames:
        wb = xl.open_workbook(file)
        try:
            spir_name = [
                i for i in wb.sheet_names() if re.match(r"^spir$", i.strip(), re.IGNORECASE)][0]
            spir_sheet = wb.sheet_by_name(spir_name)
            spirname = _get_sheet_dimension(spir_sheet)['spirname']
            max_row_col = _get_sheet_dimension(spir_sheet)['maxrow'], _get_sheet_dimension(spir_sheet)['maxcol']
            material_row_col = _find_mm_column(spir_sheet)
        except IndexError:
            spirname = 'NoSpirSheet'
            max_row_col = ('NoSpirSheet', 'NoSpirSheet')
            material_row_col = ('NoSpirSheet', 'NoSpirSheet')

        report_header = [
            "FileName", "SpirSheet", "LastCellRow", "LastCellCol", "MaterialRow", "MaterialCol",
            "Link"
        ]
        report_list.append(
            [
                os.path.split(file)[1], spirname, max_row_col[0], max_row_col[1],
                material_row_col[0], material_row_col[1], file
            ]
        )

    df = pd.DataFrame(data=report_list, columns=report_header)
    if use_relative_path:
        df['Link'] = df['Link'].apply(lambda x: os.path.relpath(x, path_to_folder))
    df['Link'] = df['Link'].apply(convert_to_hyperlink)
    df.index += 1

    # Create a Pandas Excel writer using XlsxWriter as the engine.
    writer = pd.ExcelWriter(
        os.path.join(path_to_folder, '__Quality Report for Updated SPIR(s)__.xlsx'),
        engine='xlsxwriter')
    # Convert the dataframe to an XlsxWriter Excel object.
    df.to_excel(writer, sheet_name='QA for SPIRs')
    # Get the xlsxwriter workbook and worksheet objects.
    workbook  = writer.book
    worksheet = writer.sheets['QA for SPIRs']
    # Add some cell formats.
    custom_hyperlink_format = workbook.add_format({
        'font_color': 'blue',
        # 'bold':       1,
        'underline':  1,
        'font_size':  12,
    })
    # Note: It isn't possible to format any cells that already have a format such
    # as the index or headers or any cells that contain dates or datetimes.
    # Set the format but not the column width.
    worksheet.set_column('H:H', None, custom_hyperlink_format)
    # Close the Pandas Excel writer and output the Excel file.
    writer.save()

    # Following indented code creates list of string that uses f string with multiple line elegantly
    # report_list.append(
    #     f"""
    # ##### {file[0]}: {file[1].name} #####
    #         SPIR Name: {spirname}
    #         Max (Row, Col): {max_row_col}
    #         MM(row, column): {material_row_col}
    # """
    # )


# if __name__ == '__main__':
#     codify_spir(r"C:\Users\SC-2883\Desktop\SPIR ENI\flat_folder\ENINO-4635430-v1-Coding-Updated-EJ301-J-MC-6300_C01-02.XLS")
    # codify_multiple_spir()
    # export_file_names(use_relative_path=True)
    # quality_assurance_check(r"C:/Users/SC-2883/Desktop/flat_spir")
