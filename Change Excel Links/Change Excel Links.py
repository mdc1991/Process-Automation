import pandas as pd 
import os
import re

from openpyxl import load_workbook
from openpyxl.utils.cell import column_index_from_string
from win32com.client import Dispatch

template = r'..\..\Link Change Template.xlsx' #Change the directory to where the Link Change Template is saved

def get_named_ranges(file):
    wb = load_workbook(file, read_only=True)
    defined_names = [dn.name for dn in wb.defined_names.definedName]
    defined_names_reference = [dn.attr_text for dn in wb.defined_names.definedName]
    wb.close()

    return defined_names, defined_names_reference

def get_named_range_value(ws, named_range, defined_names, defined_name_references):
    """This function returns the value from a named range which is a singular cell"""

    list_index = defined_names.index(named_range)
    excel_formula = defined_name_references[list_index]

    #get the location of the named range in terms of excel rows and columns
    split = excel_formula.split('$')
    cell_row = int(split[-1]) - 1
    cell_col = column_index_from_string(split[1]) - 1

    #return value of that cell using row and col from above
    cell_value = ws.iloc[cell_row, cell_col]
    return cell_value

def get_named_range_df(ws, named_range, defined_names, defined_names_reference):
    """This function returns the values of a named range which is a dataframe"""

    list_index = defined_names.index(named_range)
    excel_formula = defined_names_reference[list_index]

    #regex to split excluding !, $, :
    split = re.split('[!$:]', excel_formula)

    tab = split[0].strip("'")
    start_col = column_index_from_string(split[2]) - 1
    end_col = column_index_from_string(split[5])
    start_row = int(split[3]) - 1
    end_row = int(split[6]) - 1

    df = ws.iloc[start_row:end_row, start_col:end_col]
    return df

#get the cell values of named ranges in the template
names_template, references_template = get_named_ranges(template)
wb_template = pd.ExcelFile(template)
ws_template = wb_template.parse('Template', index=False, header=None)

prev_file = get_named_range_value(ws_template, 'prev_file', names_template, references_template)
new_file = get_named_range_value(ws_template, 'new_file', names_template, references_template)
links_to_change_template = get_named_range_df(ws_template, 'LinkstoChange', names_template, references_template)

#get the cell values of named ranges in the previous file
names_file, references_file = get_named_ranges(prev_file)
wb_file = pd.ExcelFile(prev_file)
ws_file = wb_file.parse('Sources of Data', index=False, header=None)
links_to_change_file = get_named_range_df(ws_file, 'LinkstoChange', names_file, references_file)  

#Check to see if the save directory exists
save_directory = '\\'.join(new_file.split('\\')[:-1])
if not os.path.exists(save_directory):
    os.makedirs(save_directory)

#Check to see if the name of the new file exists, if so delete and save over with the new file
if os.path.isfile(new_file):
    print("Removing existing file for " + new_file.split('\\')[-1].split('.xl')[0])
    os.remove(new_file)

links_to_update = links_to_change_file.iloc[:, 0].tolist()
template_links = links_to_change_template.iloc[:, 0].tolist()

xl_app = Dispatch("Excel.Application")
xl_app.Visible = False
xl_app.AskToUpdateLinks = False
xl_app.DisplayAlerts = False
xl_app.EnableEvents = False

update_wb = xl_app.Workbooks.Open(prev_file)

for link in links_to_update:
    
    prev_link = links_to_change_file[links_to_change_file.iloc[:, 0] == link].iloc[:, 1].item()

    if link in template_links:
        new_link = links_to_change_template[links_to_change_template.iloc[:, 0] == link].iloc[:, 1].item()
        
        #Check to see if new link exists
        if not os.path.isfile(new_link):
            print("Warning: " + new_link.split('\\')[-1].split('.xl')[0] + " does not exist in the directory")
        else: print("Changing links for " + link)
        
        try: 
            update_wb.ChangeLink(prev_link, new_link)
        except:
            print("Error when updating: " + link)
            
    else:
        print("Inconsistency in naming of row variables for " + link)


xl_app.AskToUpdateLinks = True
xl_app.DisplayAlerts = True
xl_app.EnableEvents = True
xl_app.Visible = True

update_wb.SaveAs(new_file)
update_wb.Close()

del xl_app
