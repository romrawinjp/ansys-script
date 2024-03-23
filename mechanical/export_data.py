#Export result
import os

#Solve the solution
# analysis = Model.Analyses[0]
# analysis.Solve(True)

#export data
name = "SILICON-9000-P1V5000"
main_path = "D:\\"
path = os.path.join(main_path, name)
if name in os.listdir(main_path):
    pass 
else: os.mkdir(path)

# Tabular Data
finished_time = '0.0001'
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

# Export Data
# for analysis in Model.Analyses:
#     results = analysis.Solution.GetChildren(DataModelObjectCategory.Result, True)
#     for result in results:
#         path = os.path.join(main_path + name, result.Name)
#         os.mkdir(path)
#         for time_step in range(1, len(T)):
#             print result.Name, "______"+str(time_step)+"_______"
#             if time_step == int(Step[-1]):
#                 result.DisplayTime = Quantity(finished_time+" [sec]")
#             else:
#                 result.DisplayTime = Quantity(T[time_step]+" [sec]")
#             result.EvaluateAllResults()
#             result.ExportToTextFile(main_path+name+"\\"+result.Name+"\\"+str(time_step)+".txt")

# def intersperse(lst, item):
#     result = [item] * (len(lst) * 2 - 1)
#     result[0::2] = lst
#     return result

# for analysis in Model.Analyses:
#     Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
#     Con = Pane.ControlUnknown
#     results = analysis.Solution.GetChildren(DataModelObjectCategory.StressProbe, True)
#     i=1
#     for result in results:
#         result.Activate()
#         path = os.path.join(main_path+name, result.Name)
#         os.mkdir(path)
#         Step = []; stress_probe = []; T = []
#         for C in range(1,4):
#             for R in range(1,Con.RowsCount+1):
#                 Text = Con.cell(R,C).Text
#                 if C == 1:
#                     Step.append(str(Text))
#                 elif C == 2:
#                     T.append(str(Text))
#                 elif C ==3 :
#                     stress_probe.append(str(Text))
#         Step = intersperse(Step, ",")
#         T = intersperse(T, ",")
#         print  "______" + result.Name +"_______"
#         stress_probe = intersperse(stress_probe, ",")
#         file = open(main_path+name+"\\" + result.Name +"\\"+str(result.Name)+".txt", "w")
#         file.writelines(Step)
#         file.write("\n")
#         file.writelines(T)
#         file.write("\n")
#         file.writelines(stress_probe)
#         file.close()
#         i+=1

# for analysis in Model.Analyses:
#     Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
#     Con = Pane.ControlUnknown
#     results = analysis.Solution.GetChildren(DataModelObjectCategory.StrainProbe, True)
#     i=1
#     for result in results:
#         result.Activate()
#         path = os.path.join(main_path+name, result.Name)
#         os.mkdir(path)
#         Step = []; stress_probe = []; T = []
#         for C in range(1,4):
#             for R in range(1,Con.RowsCount+1):
#                 Text = Con.cell(R,C).Text
#                 if C == 1:
#                     Step.append(str(Text))
#                 elif C == 2:
#                     T.append(str(Text))
#                 elif C ==3 :
#                     stress_probe.append(str(Text))
#         Step = intersperse(Step, ",")
#         T = intersperse(T, ",")
#         print  "______" + result.Name +"_______"
#         stress_probe = intersperse(stress_probe, ",")
#         file = open(main_path+name+"\\" + result.Name +"\\"+str(result.Name)+".txt", "w")
#         file.writelines(Step)
#         file.write("\n")
#         file.writelines(T)
#         file.write("\n")
#         file.writelines(stress_probe)
#         file.close()
#         i+=1
        
# for analysis in Model.Analyses:
#     Pane=ExtAPI.UserInterface.GetPane(MechanicalPanelEnum.TabularData)
#     Con = Pane.ControlUnknown
#     results = analysis.Solution.GetChildren(DataModelObjectCategory.ForceReaction, True)
#     i=1
#     for result in results:
#         result.Activate()
#         path = os.path.join(main_path+name, result.Name)
#         os.mkdir(path)
#         Step = []; linex = []; liney = []; linez = []; linet = []; T = []
#         for C in range(1,7):
#             for R in range(1,Con.RowsCount+1):
#                 Text = Con.cell(R,C).Text
#                 if C == 1:
#                     Step.append(str(Text))
#                 elif C == 2:
#                     T.append(str(Text))
#                 elif C == 3 :
#                     linex.append(str(Text))
#                 elif C == 4 :
#                     liney.append(str(Text))
#                 elif C == 5 :
#                     linez.append(str(Text))
#                 elif C == 6 :
#                     linet.append(str(Text))
#         Step = intersperse(Step, ",")
#         T = intersperse(T, ",")
#         print  "______" + result.Name +"_______"
#         linex = intersperse(linex, ",")
#         liney = intersperse(liney, ",")
#         linez = intersperse(linez, ",")
#         linet = intersperse(linet, ",")
#         file = open(main_path+name+"\\" + result.Name +"\\"+str(result.Name)+".txt", "w")
#         file.writelines(Step)
#         file.write("\n")
#         file.writelines(T)
#         file.write("\n")
#         file.writelines(linex)
#         file.write("\n")
#         file.writelines(liney)
#         file.write("\n")
#         file.writelines(linez)
#         file.write("\n")
#         file.writelines(linet)
#         file.close()
#         i+=1

# Tabular Data
finished_time = '0.0001'
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
            
print "User Defined Result"
for analysis in Model.Analyses:
    results = analysis.Solution.GetChildren(DataModelObjectCategory.UserDefinedResult, True)
    for result in results:
        print "_______" + result.Name + "_______"
        path = os.path.join(main_path+name+"\\", result.Name)
        os.mkdir(path)
        for time_step in range(1, len(T)):
            print result.Name, "______"+str(time_step)+"_______" 
            if time_step == int(Step[-1]):
                result.DisplayTime = Quantity(finished_time+" [sec]")
            else:
                result.DisplayTime = Quantity(T[time_step]+" [sec]")
            result.EvaluateAllResults()
            result.ExportToTextFile("main_path+ name +"\\"+result.Name+"\\"+str(time_step)+".txt")
