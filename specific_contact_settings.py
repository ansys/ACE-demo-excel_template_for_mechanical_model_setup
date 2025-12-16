ExtAPI = None; Ansys = None
from module_base  import *   #This will import the Ansys.Mechanical.DataModel.Enums items
import helpers
import generic_contact_settings

def Initialize(MyExtAPI, MyAnsys, my_unit_system):
    global ExtAPI; global Ansys; global unit_system
    ExtAPI = MyExtAPI; Ansys = MyAnsys ; unit_system = my_unit_system

ContactTypeInstance = generic_contact_settings.ContactTypeEnum()
ContactBehaviorInstance = generic_contact_settings.ContactBehaviorEnum()
ContactFormulationInstance = generic_contact_settings.ContactFormulationEnum()
NormalStiffnessInstance = generic_contact_settings.NormalStiffnessEnum()
InterfaceStiffnessInstance = generic_contact_settings.InterfaceStiffnessInstanceEnum()
PenetrationToleranceInstance = generic_contact_settings.PenetrationToleranceEnum()
ContactUpdateStiffnessInstance = generic_contact_settings.ContactUpdateStiffnessEnum()

def define_contact_type(contact, sp):
    contact_type = sp['Contact type']
    if contact_type.upper() in [attr for attr, value in ContactTypeInstance.__dict__.items()]:
        contact.ContactType = eval(getattr(ContactTypeInstance, contact_type.upper()))

def define_contact_behavior(contact, sp):
    contact_behavior = sp['Behavior']
    if contact_behavior.upper() in [attr for attr, value in ContactBehaviorInstance.__dict__.items()]:
        contact.Behavior = eval(getattr(ContactBehaviorInstance, contact_behavior.upper()))

def define_contact_formulation(contact, sp):
    contact_formulation = sp['Formulation']
    if contact_formulation.upper() in [attr for attr, value in ContactFormulationInstance.__dict__.items()]:
        contact.ContactFormulation = eval(getattr(ContactFormulationInstance, contact_formulation.upper()))

def define_contact_keyopt_and_opening_stiffness(contact, sp):
    contact_keyopts = sp['Keyopt settings']
    contact_opening_stiffness = sp['Contact Opening Stiffness']
    if contact_keyopts or contact_opening_stiffness:
        msg = Ansys.Mechanical.Application.Message('Changing keyoptions and real constants is done through MAPDL command snippet. Currently the code assumes CONTACT 174 elements are used.', MessageSeverityType.Warning)
        ExtAPI.Application.Messages.Add(msg)
        with Transaction():
            if contact_keyopts != '/':
                keyopt_vals = contact_keyopts.split(';')
                for child in contact.Children or []:
                    child.Delete()
                snippet = contact.AddCommandSnippet()
                for keyopt_val in keyopt_vals:
                    snippet.AppendText('keyopt,cid,'+ keyopt_val.strip('()') + '\n')
                snippet.AppendText('rmodif,1,11,'+ str(contact_opening_stiffness) + '\n')

def define_contact_normal_stiffness(contact, sp):
    contact_normal_stiffness = sp['Normal Stiffness']
    if contact_normal_stiffness.upper() in [attr for attr, value in NormalStiffnessInstance.__dict__.items()]:
        contact.NormalStiffnessValueType = eval(getattr(NormalStiffnessInstance, contact_normal_stiffness.upper()))

def define_contact_normal_stiffness_value(contact, sp):
    if sp['Normal Stiffness'] == 'Factor' or sp['Normal Stiffness'] == 'Value':
        value = sp['Normal Stiffness Value or Factor']
        contact.NormalStiffnessValue = helpers.create_quantity(value,unit_system['stiffness_unit'])

def define_contact_interface_treatment(contact, sp):
    contact_interface_treatment = sp['Interface Treatment']
    if contact_interface_treatment.upper() in [attr for attr, value in InterfaceStiffnessInstance.__dict__.items()]:
        contact.InterfaceTreatment = eval(getattr(InterfaceStiffnessInstance, contact_interface_treatment.upper()))

def define_contact_normal_offset(contact, sp):
    if sp['Interface Treatment']=='Add_Offset_Ramped_Effect' or sp['Interface Treatment']=='Add_Offset_No_Ramping':
        value = sp['Offset']
        contact.UserOffset = helpers.create_quantity(value,unit_system['length_unit'])

def define_contact_penetration_tolerance(contact, sp):
    contact_penetration_tolerance = sp['Penetration tolerance']
    if contact_penetration_tolerance.upper() in [attr for attr, value in PenetrationToleranceInstance.__dict__.items()]:
        contact.PenetrationTolerance = eval(getattr(PenetrationToleranceInstance, contact_penetration_tolerance.upper()))

def define_contact_penetration_tolerance_value(contact, sp):
    if sp['Penetration tolerance']== 'Value':
        value = sp['Penetration Tolerance Value or Factor']
        contact.PenetrationToleranceValue = helpers.create_quantity(value,unit_system['length_unit'])
    if sp['Penetration tolerance']== 'Factor':
        value = sp['Penetration Tolerance Value or Factor']
        if 0 <= value <= 1:
            contact.PenetrationToleranceFactor = value
        else: 
            msg = Ansys.Mechanical.Application.Message('Contact penetration factor must be between 0 and 1.', MessageSeverityType.Error)
            ExtAPI.Application.Messages.Add(msg)

def define_contact_update_stiffness(contact, sp):
    contact_update_stiffness = sp['Update Stiffness']
    if contact_update_stiffness.upper() in [attr for attr, value in ContactUpdateStiffnessInstance.__dict__.items()]:
        contact.UpdateStiffness = eval(getattr(ContactUpdateStiffnessInstance, contact_update_stiffness.upper()))

def SetSpecificContactSettings(analysis, workbook):
    inputdata=workbook.Worksheets("Specific Contact Settings").Select()# Select Worksheet by Name 
    ws3 = workbook.Worksheets("Specific Contact Settings")
    # Read and store values
    last_row = ws3.UsedRange.Rows.Count
    last_col = 0
    for col in range(1, ws3.UsedRange.Columns.Count + 1):
        if ws3.Cells(1, col).Value2 is None or str(ws3.Cells(1, col).Value2).strip() == "":
            break
        last_col = col
    non_empty_rows = 0
    specific_contacts = []  # list of dictionaries, one per row

    # loop over data rows (start from 2, since row 1 is header)
    for row in range(2, last_row + 1):
        row_has_data = False
        row_dict = {}
        
        for col in range(1, last_col + 1):
            key = ws3.Cells(1, col).Value2   # header
            val = ws3.Cells(row, col).Value2 # row value

            if val is not None and str(val).strip() != "":
                row_has_data = True

            row_dict[key] = val

        if row_has_data:
            specific_contacts.append(row_dict)

    # define specific contact settings
    contacts = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion)
    for sp in specific_contacts:
        contacts_with_matching_name = [c for c in contacts if sp['Contact Name Contains'] in c.Name ]
        with Transaction():
            for contact in contacts_with_matching_name:
                define_contact_type(contact, sp)
                define_contact_behavior(contact, sp)
                define_contact_formulation(contact, sp)
                if sp['Contact type'] == 'Frictional': 
                    contact.FrictionCoefficient = sp['Friction coefficient'] 
                define_contact_keyopt_and_opening_stiffness(contact, sp)
                define_contact_normal_stiffness(contact, sp)
                define_contact_normal_stiffness_value(contact, sp)
                if sp['Contact type'] == 'Rough' or sp['Contact type'] == 'Frictional' or sp['Contact type'] == 'Frictionless':
                    define_contact_interface_treatment(contact, sp)
                    define_contact_normal_offset(contact, sp)
                define_contact_penetration_tolerance(contact, sp)
                define_contact_penetration_tolerance_value(contact, sp)
                define_contact_update_stiffness(contact, sp)
    ExtAPI.DataModel.Tree.Refresh()
    return True



