# ACE-demo-excel_template_for_mechanical_model_setup
<span style="background-color: #004B87; color: white; padding: 2px 6px; border-radius: 3px; font-weight: bold;">ACE</span> | Using an Excel file to modify a model setup in Mechanical:

This repository provides the code to a demo of using an Excel file to modify a model setup in Mechanical. Several features are exposed:
- Generic contact settings: The values defined in this tab will modify all contacts of specific type. Ie, all "bonded" contacts will receive the settings provided in line 2 of this Excel tab.
  
  <img width="637" height="240" alt="image" src="https://github.com/user-attachments/assets/5cd53ee0-2970-4733-bbbb-c89365912b71" />
- Specific contact settings: Contacts which name contains the strings entered in column A of this tab receive the settings of this line.
  <img width="736" height="211" alt="image" src="https://github.com/user-attachments/assets/e74a848e-2166-4677-8e13-eb10cf3bd337" />
- Analysis settings, step controls: This enables to provide number of steps and settings for each step.
  <img width="709" height="192" alt="image" src="https://github.com/user-attachments/assets/57608b73-3826-452c-b935-4e61f7a3258f" />
- Analysis settings, other controls: This provides definition of analysis settings that are not step dependent.
<img width="577" height="382" alt="image" src="https://github.com/user-attachments/assets/6aaa3dce-5eec-4617-811f-9e9228dc89bd" />

## Important notes:
- Unsupported Content: All scripts are provided as-is and are not officially supported. Users are encouraged to review and adapt the code to suit their specific needs.
- Version-Specific Validation: Scripts have been verified to work with Ansys 2025R2. Compatibility with any other Ansys version is not guaranteed.
- No Maintenance Commitment: No ongoing maintenance, updates, or bug fixes will be provided
- Licensing and Usage: Refer to LICENSES.txt file

##  Additional information and known limitations:
  - This Excel file and the associated script are shared as an example, and should only be used as a source of inspiration to develop your own template system.
  - There is no verification of compatibility between the different options being set. Incompatible options will be ignored without any warning message being returned.
  - For specific contact settings, the lines are read from top to bottom. If there are some contacts which names contains several entries in the sheet, the settings of the last line will be applied.
  - For analysis settings, additional steps can be added by adding columns.
  - For analysis settings, leave no cell empty. If a cell is not applicable, use "N.A".
  - For force and displacement convergence, entering '0' value will set 'Calculated by solver'.
  - Make sure to change the unit system if needed so that the values entered are understood in the appropriate unit.

##  How to use this code:
  - Place all files in a same folder.
  - In Mechanical, open Mechanical scripting console and open the main file (ie, template_model_setup_main.py), and run this file. 
  - All other Python files are modules that will be imported/used when the main code is ran.
  - Running the main file from Mechanical will open a file browser for the user to select the Excel file that they want to use.
  - The xlsx file can be actually placed elsewhere to the code, to modify which folder the file browser opens by default, modify line 29 of main.py.
