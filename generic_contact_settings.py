ExtAPI = None; Ansys = None
from module_base  import *   #This will import the Ansys.Mechanical.DataModel.Enums items
import helpers

def Initialize(MyExtAPI, MyAnsys, my_unit_system):
    global ExtAPI; global Ansys; global unit_system
    ExtAPI = MyExtAPI; Ansys = MyAnsys ; unit_system = my_unit_system

class ContactTypeEnum():
    def __init__(self):
         self.BONDED = 'ContactType.Bonded'
         self.NO_SEPARATION = 'ContactType.NoSeparation'
         self.FRICTIONLESS = 'ContactType.Frictionless'
         self.ROUGH = 'ContactType.Rough'
         self.FRICTIONAL = 'ContactType.Frictional'
ContactTypeInstance = ContactTypeEnum()

class ContactBehaviorEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ContactBehavior.ProgramControlled'
         self.ASYMMETRIC = 'ContactBehavior.Asymmetric' 
         self.SYMMETRIC = 'ContactBehavior.Symmetric' 
         self.AUTOASYMMETRIC = 'ContactBehavior.AutoAsymmetric' 
ContactBehaviorInstance = ContactBehaviorEnum()

class ContactFormulationEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ContactFormulation.ProgramControlled'
         self.AUGMENTED_LAGRANGE = 'ContactFormulation.AugmentedLagrange'
         self.PURE_PENALTY = 'ContactFormulation.PurePenalty'
         self.MPC = 'ContactFormulation.MPC'
         self.NORMAL_LAGRANGE = 'ContactFormulation.NormalLagrange'
         self.BEAM = 'ContactFormulation.Beam'
ContactFormulationInstance = ContactFormulationEnum()

class NormalStiffnessEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ElementControlsNormalStiffnessType.FromContactRegion'
         self.FACTOR = 'ElementControlsNormalStiffnessType.Factor'
         self.ABSOLUTE_VALUE = 'ElementControlsNormalStiffnessType.AbsoluteValue'
NormalStiffnessInstance = NormalStiffnessEnum()

class InterfaceStiffnessInstanceEnum():
    def __init__(self):
         self.ADJUST_TO_TOUCH = 'ContactInitialEffect.AdjustToTouch'
         self.ADD_OFFSET_RAMPED_EFFECT = 'ContactInitialEffect.AddOffsetRampedEffects'
         self.ADD_OFFSET_NO_RAMPING = 'ContactInitialEffect.AddOffsetNoRamping'
InterfaceStiffnessInstance = InterfaceStiffnessInstanceEnum()

class PenetrationToleranceEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ContactPenetrationToleranceType.ProgramControlled'
         self.FACTOR = 'ContactPenetrationToleranceType.Factor'
         self.VALUE = 'ContactPenetrationToleranceType.Value'
PenetrationToleranceInstance = PenetrationToleranceEnum()

class ContactUpdateStiffnessEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'UpdateContactStiffness.ProgramControlled'
         self.NEVER = 'UpdateContactStiffness.Never'
         self.EACH_ITERATION = 'UpdateContactStiffness.EachIteration'
         self.EACH_ITERATION_AGGRESSIVE = 'UpdateContactStiffness.EachIterationAggressive'
         self.EACH_ITERATION_EXPONENTIAL = 'UpdateContactStiffness.EachIterationExponential'
ContactUpdateStiffnessInstance = ContactUpdateStiffnessEnum()

def define_contact_behavior(contact, behavior):
    contact_behavior = behavior[str(contact.ContactType)]
    if contact_behavior.upper() in [attr for attr, value in ContactBehaviorInstance.__dict__.items()]:
        contact.Behavior = eval(getattr(ContactBehaviorInstance, contact_behavior.upper()))

def define_contact_formulation(contact, formulation):
    contact_formulation = formulation[str(contact.ContactType)]
    if contact_formulation.upper() in [attr for attr, value in ContactFormulationInstance.__dict__.items()]:
        contact.ContactFormulation = eval(getattr(ContactFormulationInstance, contact_formulation.upper()))
        
def define_contact_keyopt_and_opening_stiffness(contact, keyopts, opening_stiffness):
    contact_keyopts = keyopts[str(contact.ContactType)]
    contact_opening_stiffness = opening_stiffness[str(contact.ContactType)]
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

def define_contact_normal_stiffness(contact, normal_stiffness):
    contact_normal_stiffness = normal_stiffness[str(contact.ContactType)]
    if contact_normal_stiffness.upper() in [attr for attr, value in NormalStiffnessInstance.__dict__.items()]:
        contact.NormalStiffnessValueType = eval(getattr(NormalStiffnessInstance, contact_normal_stiffness.upper()))

def define_contact_interface_treatment(contact, interface_treatment):
    contact_interface_treatment = interface_treatment[str(contact.ContactType)]
    if contact_interface_treatment.upper() in [attr for attr, value in InterfaceStiffnessInstance.__dict__.items()]:
        contact.InterfaceTreatment = eval(getattr(InterfaceStiffnessInstance, contact_interface_treatment.upper()))

def define_contact_penetration_tolerance(contact, penetration_tolerance):
    contact_penetration_tolerance = penetration_tolerance[str(contact.ContactType)]
    if contact_penetration_tolerance.upper() in [attr for attr, value in PenetrationToleranceInstance.__dict__.items()]:
        contact.PenetrationTolerance = eval(getattr(PenetrationToleranceInstance, contact_penetration_tolerance.upper()))

def define_contact_update_stiffness(contact, update_stiffness):
    contact_update_stiffness = update_stiffness[str(contact.ContactType)]
    if contact_update_stiffness.upper() in [attr for attr, value in ContactUpdateStiffnessInstance.__dict__.items()]:
        contact.UpdateStiffness = eval(getattr(ContactUpdateStiffnessInstance, contact_update_stiffness.upper()))

def SetGenericContactSettings(analysis, workbook):
    inputdata=workbook.Worksheets("Generic Contact Settings").Select()# Select Worksheet by Name 
    ws2 = workbook.Worksheets("Generic Contact Settings")
    # Read and store values
    behavior = {
        ws2.Range("A2").Value2 : ws2.Range("B2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("B3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("B4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("B5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("B6").Value2
        }
    formulation = {
        ws2.Range("A2").Value2 : ws2.Range("C2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("C3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("C4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("C5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("C6").Value2
        }
    friction_coefficient = {
        ws2.Range("A6").Value2 : ws2.Range("D6").Value2
    }
    keyopts = {
        ws2.Range("A2").Value2 : ws2.Range("E2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("E3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("E4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("E5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("E6").Value2
        }
    normal_stiffness = {
        ws2.Range("A2").Value2 : ws2.Range("F2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("F3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("F4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("F5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("F6").Value2
        }
    normal_stiffness_value = {
        ws2.Range("A2").Value2 : ws2.Range("G2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("G3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("G4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("G5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("G6").Value2
        }
    interface_treatment = {
        ws2.Range("A4").Value2 : ws2.Range("H4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("H5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("H6").Value2
        }
    interface_treatment_offset_value = {
        ws2.Range("A4").Value2 : ws2.Range("I4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("I5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("I6").Value2
        }
    penetration_tolerance = {
        ws2.Range("A2").Value2 : ws2.Range("J2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("J3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("J4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("J5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("J6").Value2
        }
    penetration_tolerance_value = {
        ws2.Range("A2").Value2 : ws2.Range("K2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("K3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("K4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("K5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("K6").Value2
        }
    opening_stiffness = {
        ws2.Range("A2").Value2 : ws2.Range("L2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("L3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("L4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("L5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("L6").Value2
        }
    update_stiffness = {
        ws2.Range("A2").Value2 : ws2.Range("M2").Value2,
        ws2.Range("A3").Value2 : ws2.Range("M3").Value2,
        ws2.Range("A4").Value2 : ws2.Range("M4").Value2,
        ws2.Range("A5").Value2 : ws2.Range("M5").Value2,
        ws2.Range("A6").Value2 : ws2.Range("M6").Value2
        }
    
    # reference contact regions
    contacts = ExtAPI.DataModel.GetObjectsByType(DataModelObjectCategory.ContactRegion)
    linear_contacts = [c for c in contacts if c.ContactType==ContactType.Bonded or c.ContactType==ContactType.NoSeparation]
    non_linear_contacts = [c for c in contacts if c not in linear_contacts]
    # Apply generic contact settings
    with Transaction():
        for contact in contacts:
            define_contact_behavior(contact, behavior)
            define_contact_formulation(contact, formulation)
            if contact.ContactType == ContactType.Frictional:
                contact.FrictionCoefficient = friction_coefficient['Frictional']    
            define_contact_keyopt_and_opening_stiffness(contact, keyopts, opening_stiffness)
            define_contact_normal_stiffness(contact, normal_stiffness)
            if contact.NormalStiffnessValueType == ElementControlsNormalStiffnessType.AbsoluteValue or contact.NormalStiffnessValueType == ElementControlsNormalStiffnessType.Factor:
                value = normal_stiffness_value[contact.ContactType.ToString()]
                contact.NormalStiffnessValue = helpers.create_quantity(value,unit_system['stiffness_unit'])
            if contact in non_linear_contacts:
                define_contact_interface_treatment(contact, interface_treatment)
                if contact.InterfaceTreatment==ContactInitialEffect.AddOffsetNoRamping or contact.InterfaceTreatment==ContactInitialEffect.AddOffsetRampedEffects:
                    value = interface_treatment_offset_value[contact.ContactType.ToString()]
                    contact.UserOffset = helpers.create_quantity(value,unit_system['length_unit'])
            define_contact_penetration_tolerance(contact, penetration_tolerance)
            if contact.PenetrationTolerance==ContactPenetrationToleranceType.Value:
                value = penetration_tolerance_value[contact.ContactType.ToString()]
                contact.PenetrationToleranceValue = helpers.create_quantity(value,unit_system['length_unit'])
            if contact.PenetrationTolerance==ContactPenetrationToleranceType.Factor:
                value = penetration_tolerance_value[contact.ContactType.ToString()]
                if 0 <= value <= 1:
                 contact.PenetrationToleranceFactor = value
                else: 
                    msg = Ansys.Mechanical.Application.Message('Contact penetration factor must be between 0 and 1.', MessageSeverityType.Error)
                    ExtAPI.Application.Messages.Add(msg)
            define_contact_update_stiffness(contact, update_stiffness)
    ExtAPI.DataModel.Tree.Refresh()
    return True
    
