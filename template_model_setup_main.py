'''
@<2025> ANSYS, Inc. Unauthorized use, distribution, or duplication is prohibited.
'''

import sys
import os
CodePath = os.getcwd() # or replace by any path you would like to use to point to the code
if not CodePath in sys.path:
    sys.path.append(CodePath)

import os
import clr
clr.AddReference("Ans.UI.Toolkit")
clr.AddReference("Microsoft.Office.Interop.Excel")
import Ansys.UI.Toolkit
import Microsoft.Office.Interop.Excel as Excel

# Import from multiple Python file
import helpers
import generic_contact_settings
import specific_contact_settings
import analysis_settings

reload(helpers)
reload(generic_contact_settings)
reload(specific_contact_settings)
reload(analysis_settings)

DefaultFolder = CodePath

filename = helpers.select_template_file(DefaultFolder)
excel = Excel.ApplicationClass()

helpers.Initialize(ExtAPI, Ansys)

workbook = excel.Workbooks.Open(filename)
unit_system = helpers.retrieve_units(workbook)

generic_contact_settings.Initialize(ExtAPI, Ansys, unit_system)
specific_contact_settings.Initialize(ExtAPI, Ansys, unit_system)
analysis_settings.Initialize(ExtAPI, Ansys, unit_system)

analysis = ExtAPI.DataModel.Project.Model.Analyses[0]

analysis_settings.SetAnalysisSettings(analysis, workbook)
generic_contact_settings.SetGenericContactSettings(analysis, workbook)
specific_contact_settings.SetSpecificContactSettings(analysis, workbook)
#excel.Application.Quit()            ## Close Only the Excel file.
#excel.Quit()                        ## Close entire Excel 
print('Finished Set Up')
