###### AUTO MESHING ########
## OBJECT: Rectangle with metal plate

import os
# Set Unit System
ExtAPI.DataModel.Project.UnitSystem = UserUnitSystemType.StandardNMM
ExtAPI.DataModel.Project.SaveProjectAfterSolution = True

# Generate Model
model = ExtAPI.DataModel.Project.Model

# Get parts
Parts = model.Geometry.GetChildren(DataModelObjectCategory.Part,True)

def preprocess():
    # Separate object and plate part in the name selection
    needle_idx = []
    plate_idx = []
    with Transaction():
        for part in Parts:
            for body in part.Children:
                idx = body.GetGeoBody().Id
                if "Solid" in body.Name:
                    needle_idx.append(body.GetGeoBody().Id)
                else: plate_idx.append(body.GetGeoBody().Id)
    print "Needle ID: ", needle_idx
    print "Body ID:", plate_idx
    
    selector = ExtAPI.SelectionManager
    selector.ClearSelection()
    selection = selector.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    
    # Declare object of interests
    obj_part = [needle_idx, plate_idx]
    obj_name = ["Solid", "Plate"]
    
    for selected_part in range(len(obj_part)):
        selection.Ids = obj_part[selected_part]
        print "Selection ID: ", selection.Ids
        # Add part to the selection
        add_part = DataModel.Project.Model.AddNamedSelection()
        add_part.Name = obj_name[selected_part]
        add_part.Location = selection
        
    # Meshing the object
    mesh_control_worksheet = Model.Mesh.Worksheet
    name_selection = Model.NamedSelections
    nsel1 = name_selection.Children[0]
    nsel2 = name_selection.Children[1]
    mesh_control_worksheet.AddRow()
    mesh_control_worksheet.SetNamedSelection(0, nsel1)
    mesh_control_worksheet.SetActiveState(0, True)
    mesh_control_worksheet.AddRow()
    mesh_control_worksheet.SetNamedSelection(0, nsel2)
    mesh_control_worksheet.SetActiveState(0, True)

# Mesh quality
def meshing(default_size, obj_size):
    mesh = Model.Mesh
    # Set Explicit Dynamics as  pre
    mesh.PhysicsPreference = MeshPhysicsPreferenceType.Explicit
    mesh.ElementSize = Quantity(default_size)
    needle_mesh = mesh.AddSizing()
    needle = ExtAPI.SelectionManager.CreateSelectionInfo(SelectionTypeEnum.GeometryEntities)
    needle.Ids = [needle_idx[0]]
    needle_mesh.ElementSize = Quantity(obj_size)
    needle_mesh.Location = needle
    mesh.GenerateMesh()
    mesh.DisplayStyle = MeshDisplayStyle.ElementQuality

preprocess()
default_size = "0.02 [mm]"
obj_size = "0.01 [mm]"
meshing(default_size, obj_size)
