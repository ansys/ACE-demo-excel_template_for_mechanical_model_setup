ExtAPI = None; Ansys = None
from module_base  import *   #This will import the Ansys.Mechanical.DataModel.Enums items
import helpers

def Initialize(MyExtAPI, MyAnsys, my_unit_system):
    global ExtAPI; global Ansys; global unit_system
    ExtAPI = MyExtAPI; Ansys = MyAnsys ; unit_system = my_unit_system

class WeakSpringsEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'WeakSpringsType.ProgramControlled'
         self.OFF = 'WeakSpringsType.Off'
         self.ON = 'WeakSpringsType.On'
WeakSpringsInstance = WeakSpringsEnum()

class SpringStiffnessEnum():
    def __init__(self):
         self.FACTOR = 'SpringsStiffnessType.Factor'
         self.MANUAL = 'SpringsStiffnessType.Manual'
         self.PROGRAM_CONTROLLED = 'SpringsStiffnessType.ProgramControlled'
SpringStiffnessInstance = SpringStiffnessEnum()

class LargeDeflectionEnum():
    def __init__(self):
         self.ON = 'True'
         self.OFF = 'False'
LargeDeflectionInstance = LargeDeflectionEnum()

class LineSearchEnum():
    def __init__(self):
         self.ON = 'LineSearchType.On'
         self.OFF = 'LineSearchType.Off'
         self.PROGRAM_CONTROLLED = 'LineSearchType.ProgramControlled'
LineSearchInstance = LineSearchEnum()

class StabilizationEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'StabilizationType.ProgramControlled'
         self.OFF = 'StabilizationType.Off'
         self.CONSTANT = 'StabilizationType.Constant'
         self.REDUCE = 'StabilizationType.Reduce'
StabilizationInstance = StabilizationEnum()

class StabilizationMethodEnum():
    def __init__(self):
         self.DAMPING = 'StabilizationMethod.Damping'
         self.ENERGY = 'StabilizationMethod.Energy'
StabilizationMethodInstance = StabilizationMethodEnum()

class StabilizationActivationForFirstSubstepEnum():
    def __init__(self):
         self.NO = 'StabilizationFirstSubstepOption.No'
         self.YES = 'StabilizationFirstSubstepOption.Yes'
         self.ONNONCONVERGENCE = 'StabilizationFirstSubstepOption.OnNonConvergence'
StabilizationActivationForFirstSubstepInstance = StabilizationActivationForFirstSubstepEnum()

class ForceConvergenceEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ConvergenceToleranceType.ProgramControlled'
         self.ON = 'ConvergenceToleranceType.On'
         self.REMOVE = 'ConvergenceToleranceType.Remove'
ForceConvergenceInstance = ForceConvergenceEnum()

class DisplacementConvergenceEnum():
    def __init__(self):
         self.PROGRAM_CONTROLLED = 'ConvergenceToleranceType.ProgramControlled'
         self.ON = 'ConvergenceToleranceType.On'
         self.REMOVE = 'ConvergenceToleranceType.Remove'
DisplacementConvergenceInstance = DisplacementConvergenceEnum()

def SetAnalysisSettings(analysis,workbook):
    # Read stepping controls
    inputdata=workbook.Worksheets("AnalysisSettings_StepControls").Select()# Select Worksheet by Name 
    ws4 = workbook.Worksheets("AnalysisSettings_StepControls")
    last_row = ws4.UsedRange.Rows.Count
    last_col = 0
    for col in range(1, ws4.UsedRange.Columns.Count + 1):
        if ws4.Cells(1, col).Value2 is None or str(ws4.Cells(1, col).Value2).strip() == "":
            break
        last_col = col
    non_empty_rows = 0
    step_definitions = {}
    for row in range(1, last_row + 1):
        row_has_data = False 
        key = ws4.Cells(row, 1).Value2   
        vals = []   
        for col in range(2, last_col + 1):             
            val = ws4.Cells(row,col).Value2 # row value
            if val is not None and str(val).strip() != "":
                row_has_data = True
                vals.append(val)
        step_definitions[key] = vals

    # Read other controls
    inputdata=workbook.Worksheets("AnalysisSettings_OtherControls").Select()
    ws5 = workbook.Worksheets("AnalysisSettings_OtherControls")
    last_row = ws5.UsedRange.Rows.Count
    other_analysis_settings = {}
    for row in range(1, last_row + 1):
        key = ws5.Cells(row, 1).Value2
        val = ws5.Cells(row,2).Value2 
        other_analysis_settings[key] = val

    ## Create analysis settings from reading Excel sheet
    analysisSettings = analysis.AnalysisSettings      
    # Apply settings
    analysisSettings.Activate()
    analysisSettings.NumberOfSteps = step_definitions['Step number'][-1]
    with Transaction():  
        # Define time stepping, per step     
        for col in range(1,last_col):
            index = col-1 # adjust index as first column is a header 
            step = int(step_definitions['Step number'][index])
            analysisSettings.CurrentStepNumber = step  
            step_end_time = helpers.create_quantity(step_definitions['Step end time'][index],unit_system['time_unit']) 
            analysisSettings.StepEndTime =  step_end_time
            if step_definitions['Autotime stepping'][index] == 'Off':
                analysisSettings.AutomaticTimeStepping = AutomaticTimeStepping.Off
                if  step_definitions['Define By'][index] == 'Substeps':
                    analysisSettings.DefineBy = TimeStepDefineByType.Substeps
                    number_of_substeps = int(float(step_definitions['Number of substeps or time of substeps'][index]))
                    analysisSettings.NumberOfSubSteps = number_of_substeps
                elif step_definitions['Define By'][index] == 'Time':
                    analysisSettings.DefineBy=TimeStepDefineByType.Time
                    number_of_substeps = helpers.create_quantity(step_definitions['Number of substeps or time of substeps'][index],unit_system['time_unit'])
                    analysisSettings.TimeStep = number_of_substeps
                else: 
                    msg = Ansys.Mechanical.Application.Message('Error in number of substep or time of substeps.', MessageSeverityType.Error)
                    ExtAPI.Application.Messages.Add(msg)
            elif step_definitions['Autotime stepping'][index] == 'On':
                analysisSettings.AutomaticTimeStepping = AutomaticTimeStepping.On
                if  step_definitions['Define By'][index] == 'Substeps':
                    initial_substep = int(float(step_definitions['Initial substeps or time'][index]))
                    minimum_substep = int(float(step_definitions['Minimum substeps or time'][index]))
                    maximum_substep = int(float(step_definitions['Maximum substeps or time'][index]))
                    analysisSettings.InitialSubsteps = initial_substep
                    analysisSettings.MinimumSubsteps = minimum_substep
                    analysisSettings.MaximumSubsteps = maximum_substep
                elif step_definitions['Define By'][index] == 'Time':
                    initial_time = helpers.create_quantity(step_definitions['Initial substeps or time'][index],unit_system['time_unit'])
                    minimum_time = helpers.create_quantity(step_definitions['Minimum substeps or time'][index],unit_system['time_unit'])
                    maximum_time = helpers.create_quantity(step_definitions['Maximum substeps or time'][index],unit_system['time_unit'])
                    analysisSettings.InitialTimeStep = initial_time
                    analysisSettings.MinimumTimeStep = minimum_time
                    analysisSettings.MaximumTimeStep = maximum_time
                else: 
                    msg = Ansys.Mechanical.Application.Message('Error in definition of initial, minimum or maximum substeps or time for automatic time stepping.', MessageSeverityType.Error)
                    ExtAPI.Application.Messages.Add(msg)
            elif step_definitions['Autotime stepping'][index] == 'Program_Controlled':
                analysisSettings.AutomaticTimeStepping = AutomaticTimeStepping.ProgramControlled
            else: 
                    msg = Ansys.Mechanical.Application.Message('Error autotime stepping definition.', MessageSeverityType.Error)
                    ExtAPI.Application.Messages.Add(msg)
        # Define other analysis settings
        # weak springs
        weak_springs = other_analysis_settings['Weak springs']
        if weak_springs.upper() in [attr for attr, value in WeakSpringsInstance.__dict__.items()]:
            analysisSettings.WeakSprings = eval(getattr(WeakSpringsInstance, weak_springs.upper()))
        # spring stiffness type
        spring_stiffness = other_analysis_settings['Spring stiffness type']
        if spring_stiffness.upper() in [attr for attr, value in SpringStiffnessInstance.__dict__.items()]:
            analysisSettings.SpringStiffness = eval(getattr(SpringStiffnessInstance, spring_stiffness.upper())) 
        # spring stiffness factor or manual
        value = float(other_analysis_settings['Spring stiffness factor or manual'])
        if analysisSettings.SpringStiffness == SpringsStiffnessType.Factor:            
            analysisSettings.SpringStiffnessFactor = value
        if analysisSettings.SpringStiffness == SpringsStiffnessType.Manual:
            analysisSettings.SpringStiffnessValue = helpers.create_quantity(value, unit_system['spring_unit'])
        # large deflection
        large_deflection = other_analysis_settings['Large deflection']
        if large_deflection.upper() in [attr for attr, value in LargeDeflectionInstance.__dict__.items()]:
            analysisSettings.LargeDeflection = eval(getattr(LargeDeflectionInstance, large_deflection.upper()))
        # line search
        line_search = other_analysis_settings['Line search']
        if line_search.upper() in [attr for attr, value in LineSearchInstance.__dict__.items()]:
            analysisSettings.LineSearch = eval(getattr(LineSearchInstance, line_search.upper()))
        # stabilization
        stabilization = other_analysis_settings['Stabilization']
        if stabilization.upper() in [attr for attr, value in StabilizationInstance.__dict__.items()]:
            analysisSettings.Stabilization = eval(getattr(StabilizationInstance, stabilization.upper()))
        # stabilization method and ratio or factor
        if analysisSettings.Stabilization == StabilizationType.Reduce or analysisSettings.Stabilization == StabilizationType.Constant:
            # stabilization method
            stabilization_method = other_analysis_settings['Stabilization method']
            if stabilization_method.upper() in [attr for attr, value in StabilizationMethodInstance.__dict__.items()]:
                analysisSettings.StabilizationMethod = eval(getattr(StabilizationMethodInstance, stabilization_method.upper()))
                # stabilization ratio or factor
                value = float(other_analysis_settings['Stabilization ratio or factor'])
                if analysisSettings.StabilizationMethod == StabilizationMethod.Damping:
                    analysisSettings.StabilizationDampingFactor = value
                if analysisSettings.StabilizationMethod == StabilizationMethod.Energy:
                    analysisSettings.StabilizationEnergyDissipationRatio = value
            # activation for first substep
            activation_for_first_substep = other_analysis_settings['Stabilization activation for first substep']
            if activation_for_first_substep.upper() in [attr for attr, value in StabilizationActivationForFirstSubstepInstance.__dict__.items()]:
                analysisSettings.StabilizationActivationForFirstSubstep = eval(getattr(StabilizationActivationForFirstSubstepInstance, activation_for_first_substep.upper()))
            # stabilization limit
            stabilization_limit = float(other_analysis_settings['Stabilization limit'])
            analysisSettings.StabilizationForceLimit = stabilization_limit
            # force convergence
            force_convergence = other_analysis_settings['Force convergence']
            if force_convergence.upper() in [attr for attr, value in ForceConvergenceInstance.__dict__.items()]:
                analysisSettings.ForceConvergence = eval(getattr(ForceConvergenceInstance, force_convergence.upper()))
            # force convergence value and percentage
            if analysisSettings.ForceConvergence == ConvergenceToleranceType.On:
                force_convergence_value = other_analysis_settings['Force convergence value']
                analysisSettings.ForceConvergenceValue = helpers.create_quantity(force_convergence_value, unit_system['force_unit'])
                force_convergence_tolerance = other_analysis_settings['Force convergence tolerance percentage']
                analysisSettings.ForceConvergenceTolerance = helpers.create_quantity(force_convergence_tolerance, unit_system['force_unit'])
            # displacement convergence
            displacement_convergence = other_analysis_settings['Displacement convergence']
            if displacement_convergence.upper() in [attr for attr, value in DisplacementConvergenceInstance.__dict__.items()]:
                analysisSettings.DisplacementConvergence = eval(getattr(DisplacementConvergenceInstance, displacement_convergence.upper()))
            # force convergence value and percentage
            if analysisSettings.DisplacementConvergence == ConvergenceToleranceType.On:
                displacement_convergence_value = other_analysis_settings['Displacement convergence value']
                analysisSettings.DisplacementConvergenceValue = helpers.create_quantity(displacement_convergence_value, unit_system['length_unit'])
                displacement_convergence_tolerance = other_analysis_settings['Displacement convergence tolerance percentage']
                analysisSettings.DisplacementConvergenceTolerance = helpers.create_quantity(displacement_convergence_tolerance, unit_system['length_unit'])





    



        


            
    ExtAPI.DataModel.Tree.Refresh()       
    return True