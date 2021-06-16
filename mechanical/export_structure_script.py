import os
# Set Unit System
ExtAPI.DataModel.Project.UnitSystem = UserUnitSystemType.StandardNMM
ExtAPI.DataModel.Project.SaveProjectAfterSolution = True

# Generate Model
model = ExtAPI.DataModel.Project.Model

'''# Assign Materials to Geometry
for i in range(len(Model.Materials.Children)):
    if "Dermis" == Model.Materials.Children[i].Name:
        Model.Geometry.Children[0].Children[0].Material = Model.Materials.Children[i].Name
    elif "Epidermis" == Model.Materials.Children[i].Name:
        Model.Geometry.Children[1].Children[0].Material = Model.Materials.Children[i].Name
    elif "Stratum Corneum" == Model.Materials.Children[i].Name:
        Model.Geometry.Children[2].Children[0].Material = Model.Materials.Children[i].Name
    elif "PU" == Model.Materials.Children[i].Name:
        Model.Geometry.Children[3].Children[0].Material = Model.Materials.Children[i].Name
'''
# Get parts
Parts = model.Geometry.GetChildren(DataModelObjectCategory.Part,True)

Skin_Ids = []
Needle_Ids = []
# print object ids each part
with Transaction():
    for part in Parts:
        for body in part.Children:
            if "Needle" in body.Name:
                Needle_Ids.append(body.GetGeoBody().Id)
            else: Skin_Ids.append(body.GetGeoBody().Id)
            print("Body name: ", body.Name)
            print("Body name: ", body.GetGeoBody().Id)
print(Skin_Ids)
print(Needle_Ids)

# Selection
SlMn = ExtAPI.SelectionManager
SlMn.ClearSelection()
Sel = SlMn.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)

p = [Skin_Ids, Needle_Ids]
pname = ["Skin", "Needle"]
for s in range(2):
    Sel.Ids = p[s]
    print("Sel.Ids: ",Sel.Ids)
    NS = DataModel.Project.Model.AddNamedSelection()
    NS.Name = pname[s]
    NS.Location = Sel

# Generate mesh
mws = Model.Mesh.Worksheet
nsels = Model.NamedSelections
ns1 = nsels.Children[0]
ns2 = nsels.Children[1]
mws.AddRow()
mws.SetNamedSelection(0,ns1)
mws.SetActiveState(0,True)
mws.AddRow()
mws.SetNamedSelection(1,ns2)
mws.SetActiveState(1,True)

# Mesh quality
mesh = Model.Mesh
mesh.ElementSize = Quantity("0.0155 [mm]")
needle_mesh = mesh.AddSizing()
needle = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
needle.Ids = [Needle_Ids[0]]
needle_mesh.ElementSize = Quantity("0.01 [mm]")
needle_mesh.Location = needle
mesh.GenerateMesh()
mesh.DisplayStyle = MeshDisplayStyle.ElementQuality

#Set Camera
camera = Graphics.Camera
camera.FocalPoint = Point((-0.5,0.1,0.0), "mm")
camera.SetFit()

# Export Image
setting2d = Ansys.Mechanical.Graphics.GraphicsImageExportSettings()
setting2d.Resolution = GraphicsResolutionType.EnhancedResolution
Graphics.ExportImage("D:\\mesh.png", GraphicsImageExportFormat.PNG, setting2d)

#Skin Contact
ns = DataModel.Project.Model.AddNamedSelection()
ns.Name = "SkinContact"
ns.ScopingMethod = GeometryDefineByType.Worksheet
GenerationCriteria = ns.GenerationCriteria
c1 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c1.Action = SelectionActionType.Add
c1.EntityType = SelectionType.GeoFace
c1.Criterion = SelectionCriterionType.LocationY
c1.Operator = SelectionOperatorType.Equal
c1.Value = Quantity("5.334e-004 [m]")
GenerationCriteria.Add(c1)
ns.Generate()

#Needle Contact
ns1 = DataModel.Project.Model.AddNamedSelection()
ns1.Name = "NeedleContact"
ns1.ScopingMethod = GeometryDefineByType.Worksheet
GenerationCriteria = ns1.GenerationCriteria
c2 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c2.Action = SelectionActionType.Add
c2.EntityType = SelectionType.GeoFace
c2.Criterion = SelectionCriterionType.LocationY
c2.Operator = SelectionOperatorType.Equal
c2.Value = Quantity("5.842e-004 [m]")
GenerationCriteria.Add(c2)
ns1.Generate()

#Fix Support
ns2 = DataModel.Project.Model.AddNamedSelection()
ns2.Name = "FixedSupport"
ns2.ScopingMethod = GeometryDefineByType.Worksheet
GenerationCriteria = ns2.GenerationCriteria
c3 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c3.Action = SelectionActionType.Add
c3.EntityType = SelectionType.GeoFace
c3.Criterion = SelectionCriterionType.LocationY
c3.Operator = SelectionOperatorType.Equal
c3.Value = Quantity("0 [m]")
GenerationCriteria.Add(c3)
c4 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c4.Action = SelectionActionType.Add
c4.EntityType = SelectionType.GeoFace
c4.Criterion = SelectionCriterionType.LocationZ
c4.Operator = SelectionOperatorType.Equal
c4.Value = Quantity("0 [m]")
GenerationCriteria.Add(c4)
c5 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c5.Action = SelectionActionType.Add
c5.EntityType = SelectionType.GeoFace
c5.Criterion = SelectionCriterionType.LocationX
c5.Operator = SelectionOperatorType.Equal
c5.Value = Quantity("0 [m]")
GenerationCriteria.Add(c5)
c6 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c6.Action = SelectionActionType.Add
c6.EntityType = SelectionType.GeoFace
c6.Criterion = SelectionCriterionType.LocationZ
c6.Operator = SelectionOperatorType.Equal
c6.Value = Quantity("3.81e-004 [m]")
GenerationCriteria.Add(c6)
c7 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c7.Action = SelectionActionType.Add
c7.EntityType = SelectionType.GeoFace
c7.Criterion = SelectionCriterionType.LocationX
c7.Operator = SelectionOperatorType.Equal
c7.Value = Quantity("-3.81e-004 [m]")
GenerationCriteria.Add(c7)
c8 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c8.Action = SelectionActionType.Add
c8.EntityType = SelectionType.GeoFace
c8.Criterion = SelectionCriterionType.LocationX
c8.Operator = SelectionOperatorType.Equal
c8.Value = Quantity("-1.1225e-007 [m]")
GenerationCriteria.Add(c8)
ns2.Generate()

#Needlesides
ns3 = DataModel.Project.Model.AddNamedSelection()
ns3.Name = "NeedleSides"
ns3.ScopingMethod = GeometryDefineByType.Worksheet
GenerationCriteria = ns3.GenerationCriteria
c9 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c9.Action = SelectionActionType.Add
c9.EntityType = SelectionType.GeoFace
c9.Criterion = SelectionCriterionType.LocationY
c9.Operator = SelectionOperatorType.GreaterThan
c9.Value = Quantity("5.842e-004 [m]")
GenerationCriteria.Add(c9)
ns3.Generate()

# Pressure
ns4 = DataModel.Project.Model.AddNamedSelection()
ns4.Name = "Pressure"
ns4.ScopingMethod = GeometryDefineByType.Worksheet
GenerationCriteria = ns4.GenerationCriteria
c11 = Ansys.ACT.Automation.Mechanical.NamedSelectionCriterion()
c11.Action = SelectionActionType.Add
c11.EntityType = SelectionType.GeoFace
c11.Criterion = SelectionCriterionType.LocationY
c11.Operator = SelectionOperatorType.GreaterThan
c11.Value = Quantity("8.89e-004 [m]")
GenerationCriteria.Add(c11)
ns4.Generate()

# Add connection
connection = ExtAPI.DataModel.Project.Model.Connections
contact_region = connection.AddContactRegion()
contact_region.Name = "Needle-Skin"
contact_region.ContactType = ContactType.Frictionless
contact_region.SourceLocation = Model.NamedSelections.Children[2]
contact_region.TargetLocation = Model.NamedSelections.Children[3]

contact_region1 = connection.Children[1].AddContactRegion()
contact_region1.ContactType = ContactType.Frictional
contact_region1.SourceLocation = Model.NamedSelections.Children[2]
contact_region1.TargetLocation = Model.NamedSelections.Children[5]
contact_region1.FrictionCoefficient = 0.42

#Simulation
# Add-Pressure
finished_time = '0.00005'
analysis = Model.Analyses[0]
pressure = analysis.AddPressure()
pressure.Location = Model.NamedSelections.Children[6]
pressure.Magnitude.Inputs[0].DiscreteValues = [Quantity("0 [s]"), Quantity("2e-5 [s]"), Quantity("2.5e-5 [s]"), Quantity(finished_time+" [s]")]
pressure.Magnitude.Output.DiscreteValues = [Quantity("300 [MPa]"), Quantity("300 [MPa]"), Quantity("0 [MPa]"), Quantity("0 [MPa]")]

fixed_support = analysis.AddFixedSupport()
fixed_support.Location = Model.NamedSelections.Children[4]

#dropheight = analysis.AddInitialVelocity()
#dropheight.Location = Model.NamedSelections.Children[1]
#dropheight.InputType = InitialConditionsType.DropHeight
#dropheight.DropHeight = Quantity("1e-003 [m]")

# Time
analysis_setting = analysis.AnalysisSettings
analysis_setting.NumberOfSteps = 4
analysis_setting.SetStepEndTime(1, Quantity("0.00001 [s]"))
analysis_setting.SetStepEndTime(2, Quantity("0.00002 [s]"))
analysis_setting.SetStepEndTime(3, Quantity("0.000025 [s]"))
analysis_setting.SetStepEndTime(4, Quantity(finished_time+" [s]"))

# Temperature
model.Analyses[0].EnvironmentTemperature = Quantity('36 [C]')

# Solution.
solution = ExtAPI.DataModel.Project.Model.Analyses[0].Solution
total_deformation = solution.AddTotalDeformation()
total_deformation = solution.AddTotalDeformation()
skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
skinp.Ids = [Needle_Ids[0]]
total_deformation.Location = skinp
for i in range(3):
    total_deformation = solution.AddTotalDeformation()
    skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    skinp.Ids = [Skin_Ids[i]]
    total_deformation.Location = skinp
total_stress = solution.AddEquivalentStress()
total_stress = solution.AddEquivalentStress()
skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
skinp.Ids = [Needle_Ids[0]]
total_stress.Location = skinp
for i in range(3):
    total_stress = solution.AddEquivalentStress()
    skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    skinp.Ids = [Skin_Ids[i]]
    total_stress.Location = skinp
total_pstrain = solution.AddEquivalentPlasticStrain()
total_pstrain = solution.AddEquivalentPlasticStrain()
skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
skinp.Ids = [Needle_Ids[0]]
total_pstrain.Location = skinp
for i in range(3):
    total_pstrain = solution.AddEquivalentPlasticStrain()
    skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    skinp.Ids = [Skin_Ids[i]]
    total_pstrain.Location = skinp
total_estrain = solution.AddEquivalentElasticStrain()
total_estrain = solution.AddEquivalentElasticStrain()
skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
skinp.Ids = [Needle_Ids[0]]
total_estrain.Location = skinp
for i in range(3):
    total_estrain = solution.AddEquivalentElasticStrain()
    skinp = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    skinp.Ids = [Skin_Ids[i]]
    total_estrain .Location = skinp

# Add User defined result
no_unit = ["MATERIAL"]
pres = ["PRESSURE"]
mass = ["MASSALL"]
density = ["DENSITY"]
dis = ["UX", "UY", "UZ", "USUM", "UVECTORS", "LOCX", "LOCY", "LOCZ", "LOC_DEFX", "LOC_DEFY", "LOC_DEFZ"]
vel = ["VX", "VY", "VZ", "VSUM", "VVECTORS"]
acc = ["AX", "AY", "AZ", "ASUM", "AVECTORS"]
stress = ["SX", "SY", "SZ", "SXY", "SYZ", "SXZ", "S1", "S2", "S3", "SINT", "SEQV", "SVECTORS", "SMAXSHEAR"]
strain = [ "EPELX", "EPELY", "EPELZ", "EPELXY", "EPELYZ", "EPELXZ", "EPEL1", "EPEL2", "EPEL3", "EPELINT", "EPELVECTORS", "EPELEQV_RST", "STRAIN_XX", "STRAIN_YY", "STRAIN_ZZ", "STRAIN_XY", "STRAIN_YZ", "STRAIN_ZX", "EPPLX", "EPPLY", "EPPLZ", "EPPLXY", "EPPLYZ", "EPPLXZ", "EPPL1", "EPPL2", "EPPL3", "EPPLINT", "EPPLVECTORS", "EPPLEQV_RST"]

user_defined_result = no_unit + pres + mass + density + dis + vel + acc + stress + strain

for udr in range(len(user_defined_result)):
    udr_result = solution.AddUserDefinedResult()
    udr_result.Name = user_defined_result[udr]
    udr_result.Expression = user_defined_result[udr]
    if user_defined_result[udr] in dis:
        udr_result.OutputUnit = UnitCategoryType.Displacement
    elif user_defined_result[udr] in no_unit:
        udr_result.OutputUnit = UnitCategoryType.NoUnits
    elif user_defined_result[udr] in pres:
        udr_result.OutputUnit = UnitCategoryType.Pressure
    elif user_defined_result[udr] in mass:
        udr_result.OutputUnit = UnitCategoryType.Mass
    elif user_defined_result[udr] in density:
        udr_result.OutputUnit = UnitCategoryType.Density
    elif user_defined_result[udr] in vel:
        udr_result.OutputUnit = UnitCategoryType.Velocity
    elif user_defined_result[udr] in acc:
        udr_result.OutputUnit = UnitCategoryType.Acceleration
    elif user_defined_result[udr] in stress:
        udr_result.OutputUnit = UnitCategoryType.Stress
    elif user_defined_result[udr] in strain:
        udr_result.OutputUnit = UnitCategoryType.Strain

probe = solution.AddStressProbe()
probe.GeometryLocation = Model.NamedSelections.Children[3]
probe.ResultSelection =  ProbeDisplayFilter.Equivalent

analysis.Solve(True)

# TabularData
finished_time = '0.00005'
solution = ExtAPI.DataModel.Project.Model.Analyses[0].Solution
solution.Activate()
Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
Con = Pane.ControlUnknown
Step = []; T = []
for C in range(1,3):
    for R in range(1,Con.RowsCount+1):
        Text = Con.cell(R,C).Text
        if C == 1:
            Step.append(str(Text))
        elif C == 2:
            T.append(str(Text))

name = "Hollow Structure"
path = os.path.join("D:\\",name)
os.mkdir(path)

# Export Data
for analysis in Model.Analyses:
    results = analysis.Solution.GetChildren(DataModelObjectCategory.Result, True)
    for result in results:
        path = os.path.join("D:\\" + name, result.Name)
        os.mkdir(path)
        for time_step in range(1, len(T)):
            print(result.Name, "______"+str(time_step)+"_______")
            if time_step == int(Step[-1]):
                result.DisplayTime = Quantity(finished_time+" [sec]")
            else:
                result.DisplayTime = Quantity(T[time_step]+" [sec]")
            result.EvaluateAllResults()
            result.ExportToTextFile("D:\\"+name+"\\"+result.Name+"\\"+str(time_step)+".txt")

for analysis in Model.Analyses:
    results = analysis.Solution.GetChildren(DataModelObjectCategory.UserDefinedResult, True)
    for result in results:
        path = os.path.join("D:\\"+name, result.Name)
        os.mkdir(path)
        for time_step in range(1, len(T)):
            print(result.Name, "______"+str(time_step)+"_______")
            if time_step == int(Step[-1]):
                result.DisplayTime = Quantity(finished_time+" [sec]")
            else:
                result.DisplayTime = Quantity(T[time_step]+" [sec]")
            result.EvaluateAllResults()
            result.ExportToTextFile("D:\\"+name+"\\"+result.Name+"\\"+str(time_step)+".txt")

def intersperse(lst, item):
    result = [item] * (len(lst) * 2 - 1)
    result[0::2] = lst
    return result

for analysis in Model.Analyses:
    Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    Con = Pane.ControlUnknown
    results = analysis.Solution.GetChildren(DataModelObjectCategory.StressProbe, True)
    i=1
    for result in results:
        path = os.path.join("D:\\"+name, result.Name)
        os.mkdir(path)
        Step = []; stress_probe = []; T = []
        for C in range(1,4):
            for R in range(1,Con.RowsCount+1):
                Text = Con.cell(R,C).Text
                if C == 1:
                    Step.append(str(Text))
                elif C == 2:
                    T.append(str(Text))
                elif C ==3 :
                    stress_probe.append(str(Text))
        Step = intersperse(Step, ",")
        T = intersperse(T, ",")
        stress_probe = intersperse(stress_probe, ",")
        file = open("D:\\"+name+"\\" + result.Name +"\\stress_probe"+str(i)+".txt", "w")
        file.writelines(Step)
        file.write("\n")
        file.writelines(T)
        file.write("\n")
        file.writelines(stress_probe)
        file.close()
        i+=1

for analysis in Model.Analyses:
    Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    Con = Pane.ControlUnknown
    results = analysis.Solution.GetChildren(DataModelObjectCategory.StrainProbe, True)
    i=1
    for result in results:
        path = os.path.join("D:\\"+name, result.Name)
        os.mkdir(path)
        Step = []; stress_probe = []; T = []
        for C in range(1,4):
            for R in range(1,Con.RowsCount+1):
                Text = Con.cell(R,C).Text
                if C == 1:
                    Step.append(str(Text))
                elif C == 2:
                    T.append(str(Text))
                elif C ==3 :
                    stress_probe.append(str(Text))
        Step = intersperse(Step, ",")
        T = intersperse(T, ",")
        stress_probe = intersperse(stress_probe, ",")
        file = open("D:\\"+name+"\\" + result.Name +"\\strain_probe"+str(i)+".txt", "w")
        file.writelines(Step)
        file.write("\n")
        file.writelines(T)
        file.write("\n")
        file.writelines(stress_probe)
        file.close()
        i+=1

for analysis in Model.Analyses:
    Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    Con = Pane.ControlUnknown
    results = analysis.Solution.GetChildren(DataModelObjectCategory.ForceReaction, True)
    i = 1
    for result in results:
        path = os.path.join("D:\\"+name, result.Name)
        os.mkdir(path)
        Step = []; stress_probe = []; T = []
        for C in range(1,4):
            for R in range(1,Con.RowsCount+1):
                Text = Con.cell(R,C).Text
                if C == 1:
                    Step.append(str(Text))
                elif C == 2:
                    T.append(str(Text))
                elif C ==3 :
                    stress_probe.append(str(Text))
        Step = intersperse(Step, ",")
        T = intersperse(T, ",")
        stress_probe = intersperse(stress_probe, ",")
        file = open("D:\\"+name+"\\" + result.Name +"\\force_reaction"+str(i)+".txt", "w")
        file.writelines(Step)
        file.write("\n")
        file.writelines(T)
        file.write("\n")
        file.writelines(stress_probe)
        file.close()
        i+=1

for analysis in Model.Analyses:
    Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
    Con = Pane.ControlUnknown
    results = analysis.Solution.GetChildren(DataModelObjectCategory.DeformationProbe, True)
    i = 1
    for result in results:
        path = os.path.join("D:\\"+name, result.Name)
        os.mkdir(path)
        Step = []; stress_probe = []; T = []
        for C in range(1,4):
            for R in range(1,Con.RowsCount+1):
                Text = Con.cell(R,C).Text
                if C == 1:
                    Step.append(str(Text))
                elif C == 2:
                    T.append(str(Text))
                elif C ==3 :
                    stress_probe.append(str(Text))
        Step = intersperse(Step, ",")
        T = intersperse(T, ",")
        stress_probe = intersperse(stress_probe, ",")
        file = open("D:\\"+name+"\\" + result.Name +"\\deformation_probe"+str(i)+".txt", "w")
        file.writelines(Step)
        file.write("\n")
        file.writelines(T)
        file.write("\n")
        file.writelines(stress_probe)
        file.close()
        i+=1

#Animation settings
Graphics.ResultAnimationOptions.NumberOfFrames = 50
Graphics.ResultAnimationOptions.Duration = Quantity(5, 's')
settings = Ansys.Mechanical.Graphics.AnimationExportSettings(width = 1000, height = 665)

path = os.path.join("D:\\"+name, "video")
os.mkdir(path)
for analysis in Model.Analyses:
    results = analysis.Solution.GetChildren(DataModelObjectCategory.Result, True)
    for result in results:
	    location ="D:\\"+name+"\\video\\" + result.Name + ".wmv"
	    result.ExportAnimation(location, GraphicsAnimationExportFormat.WMV, settings)

