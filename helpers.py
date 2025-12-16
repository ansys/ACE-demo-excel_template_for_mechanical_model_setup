ExtAPI = None; Ansys = None
from module_base  import *   #This will import the Ansys.Mechanical.DataModel.Enums items

def Initialize(MyExtAPI, MyAnsys):
    global ExtAPI; global Ansys
    ExtAPI = MyExtAPI; Ansys = MyAnsys

def select_template_file(DefaultFolder):
    FilePath =Ansys.UI.Toolkit.FileDialog.ShowOpenFilesDialog(Ansys.UI.Toolkit.Dialog(),DefaultFolder,'Excel Files(s)|*.xls*',0,'Select .xlsx file',None)
    if str(FilePath[0]) != 'OK':
        return None
    filename = list(FilePath[1])[0] # only one file should be selected
    return filename

def retrieve_units(workbook):
    inputdata=workbook.Worksheets("Unit System").Select()# Select Worksheet by Name 
    ws2 = workbook.Worksheets("Unit System")
    unit_system_string = ws2.Range("B1").Value2
    units = unit_system_string.split(",")
    length_unit = units[0]
    weight_unit = units[1]
    force_unit = units[2]
    time_unit = units[3]
    stiffness_unit = force_unit + ' ' + length_unit + '^-1 ' + length_unit + '^-1 ' + length_unit + '^-1 '
    spring_unit = force_unit + ' ' + length_unit + '^-1 '
    unit_system = {
      'force_unit' :   force_unit,
      'length_unit' : length_unit,
      'stiffness_unit' : stiffness_unit,
      'time_unit' : time_unit,
      'spring_unit' : spring_unit
    }
    return unit_system

def create_quantity(value,unit):    
    return Quantity(str(value) + ' [' + unit + ']')