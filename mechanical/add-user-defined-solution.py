solution = ExtAPI.DataModel.Project.Model.Analyses[0].Solution
dis = ["UX", "UY", "UZ", "USUM", "UVECTORS", "LOCX", "LOCY", "LOCZ", "LOC_DEFX", "LOC_DEFY", "LOC_DEFZ"]
user_defined_result = dis

name = "LDPE1-5000"

# Tabular Data
finished_time = '0.001'
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

for udr in range(len(user_defined_result)):
    udr_result = solution.AddUserDefinedResult()
    udr_result.Name = user_defined_result[udr]
    udr_result.Expression = user_defined_result[udr]
    if user_defined_result[udr] in dis:
        udr_result.OutputUnit = UnitCategoryType.Displacement
