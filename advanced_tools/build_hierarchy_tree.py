import os
import csv
import xlrd
from tkinter import messagebox, filedialog, Tk
from advanced_tools import hierarchy_tree, outlined_hierarchy, read_excel_to_lol

def main():
    # fname = r"C:\Users\sc-2883\Desktop\SAP FL Extract.xlsx"
    window = Tk()
    window.wm_withdraw()

    messagebox.showinfo(title="Open File",
                        message="Welcome! Please Select Tag, Description, Parent tag CSV file",
                        detail="Press OK to continue")

    fname = os.path.abspath(
        filedialog.askopenfilename(filetypes=(("CSV files", "*.xls;*.xlsx;*.csv"),
                                            ("All files", "*.*")))
        )

    output_folder = os.path.splitext(fname)[0]

    if os.path.isdir(output_folder):
        pass
    else:
        os.mkdir(output_folder)

    # Hierarchy building starts here
    if 'xls' in os.path.splitext(fname)[1]:
        csv_content = read_excel_to_lol(fname)
    else:
        with open(fname, 'r', encoding='latin1') as csvfile:
            reader = csv.reader(csvfile, delimiter=';')
            csv_content = list(reader)

    # Fill empty Description Cell
    # Fill empty parent tag field with tags itself to make it 1st level tag
    for row in csv_content:
        row[1] = row[1].replace('\n', '')
        if row[1] == '':
            row[1] = ">No_Description_for_{}".format(row[0])
        if row[2] == '':
            row[2] = row[0]

    tags = set([row[0] for row in csv_content[1:]])
    parents = set([row[2] for row in csv_content[1:]])
    for i in parents.difference(tags):
        csv_content.append([i, '', i, ''])

    tag_desc_dict = dict(zip([i[0] for i in csv_content], [i[1] for i in csv_content]))

    table = [(i[0]+'__'+i[1], i[2]+'__'+tag_desc_dict[i[2]]) for i in csv_content[1:]]

    # Setting output files with paths
    output_wo_ext = os.path.splitext(os.path.split(fname)[1])[0]
    txt_output = os.path.join(output_folder, output_wo_ext + '_outlined.txt')
    excel_output = os.path.join(output_folder, output_wo_ext + '_outlined.xlsx')

    # Create tree-like hierarchical tag render in txt file
    hierarchy_tree(table=table, output_file=txt_output, on_screen=False)

    # Writing into excel with outline structure, after reading hierarchy_tree txt file
    outlined_hierarchy(txtfile=txt_output, sysname="System Name:", sysno="System Number:",
                    wkbk=excel_output, ws="Tag_Hierarchy")

    messagebox.showinfo(title="Complete.",
                        message="It is all done! Thanks for using.",
                        detail="Please check the same folder for hierarchy file.")

