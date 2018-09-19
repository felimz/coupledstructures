import os
import sys
import sap2000
from collections import OrderedDict
from modelclasses import Props

# %% OPEN SAP2000 AND READY PROGRAM

# check working directory
api_path_home = r'C:\Users\Felipe\Dropbox\PycharmProjects\CoupledStructures\models'
api_path_fluor = r'C:\Users\mej12981\PycharmProjects\CoupledStructures\models'
try:
    os.chdir(api_path_home)
except FileNotFoundError:
    try:
        os.chdir(api_path_fluor)
    except FileNotFoundError:
        print("Could not find api_path directory")
        sys.exit(1)
    else:
        api_path = api_path_fluor
else:
    api_path = api_path_home
    sap2000.checkinstall(api_path)

# create sap2000 object in memory
attach_to_instance = False
specify_path = False  # program finds most recent installation after SAP2000 v19. For other installs, specify path.
program_path = r'C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe'
my_sap_object = sap2000.attachtoapi(attach_to_instance, specify_path, program_path)

# open sap2000
visible = False
sap_model = sap2000.opensap2000(my_sap_object, visible)

# open new model in units of kip_in_F
sap2000.newmodel(sap_model)

#%% DEFINE MODEL GEOMETRY, PROPERTIES, AND LOADING

# initialize frame properties object
props_obj = Props()

# set degrees of freedom for 2-D motion into sap2000
props_obj.load_mdl_dof_df(sap_model, '2-D')

# load dataframe of material properties into sap2000
props_obj.load_mat_props_df(sap_model)

# load dataframe of frame properties into sap2000
props_obj.load_frm_props_df(sap_model)

print('hello world')

# define frame section property modifiers

ModValue = [1, 1, 1, 1, 1, 1, 1, 1]

sap_model.PropFrame.SetModifiers('R1', ModValue)

# switch to k-ft units

sap_model.SetPresentUnits(sap2000.UNITS['kip_ft_F'])

# add frame object by coordinates

[Frame1, ret] = sap_model.FrameObj.AddByCoord(0, 0, 0, 0, 0, 20,  '', 'R1', '', 'Global')
[Frame2, ret] = sap_model.FrameObj.AddByCoord(20, 0, 0, 20, 0, 20, '', 'R1', '', 'Global')
[Frame3, ret] = sap_model.FrameObj.AddByCoord(0, 0, 20, 20, 0, 20, '', 'R2', '', 'Global')
[Frame4, ret] = sap_model.FrameObj.AddByCoord(40, 0, 0, 40, 0, 20, '', 'R3', '', 'Global')
[Frame5, ret] = sap_model.FrameObj.AddByCoord(60, 0, 0, 60, 0, 20, '', 'R3', '', 'Global')
[Frame6, ret] = sap_model.FrameObj.AddByCoord(40, 0, 20, 60, 0, 20, '', 'R4', '', 'Global')


# assign point object restraint at base

Restraint = [True, True, True, True, True, True]

[Frame1i, Frame1j, ret] = sap_model.FrameObj.GetPoints(Frame1, '', '')

sap_model.PointObj.SetRestraint(Frame1i, Restraint)

[Frame2i, Frame1j, ret] = sap_model.FrameObj.GetPoints(Frame2, '', '')

sap_model.PointObj.SetRestraint(Frame2i, Restraint)

[Frame4i, Frame4j, ret] = sap_model.FrameObj.GetPoints(Frame4, '', '')

sap_model.PointObj.SetRestraint(Frame4i, Restraint)

[Frame5i, Frame5j, ret] = sap_model.FrameObj.GetPoints(Frame5, '', '')

sap_model.PointObj.SetRestraint(Frame5i, Restraint)

# refresh view, update (initialize) zoom

sap_model.View.RefreshView(0, False)

# add load patterns

load_patterns = {1: 'LTYPE_DEAD', 2: 'LTYPE_OTHER'}

sap_model.LoadPatterns.Add(load_patterns[1], sap2000.LOAD_PATTERN_TYPES[load_patterns[1]], 0, True)
sap_model.LoadPatterns.Add(load_patterns[2], sap2000.LOAD_PATTERN_TYPES[load_patterns[2]], 0, True)

# assign loading for load pattern 'D'

[Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')
sap_model.PointObj.SetLoadForce(Frame3i, load_patterns[1], [0, 0, -10, 0, 0, 0])
sap_model.PointObj.SetLoadForce(Frame3j, load_patterns[1], [0, 0, -10, 0, 0, 0])

# assign loading for load pattern 'EQ'
[Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')
sap_model.FrameObj.SetLoadDistributed(Frame3, load_patterns[2], 1, 10, 0, 1, 1.8, 1.8)

# switch to k-in units

sap_model.SetPresentUnits(sap2000.UNITS['kip_in_F'])

#%% SAVE MODEL AND RUN IT

file_name = 'API_1-001.sdb'
sap2000.saveandrunmodel(sap_model, api_path, file_name)

#%% OBTAIN RESULTS FROM SAP2000 MODEL

# initialize for Sap2000 results

[Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')

# get Sap2000 results for all load patterns:

SapResult = OrderedDict()

for key, val in load_patterns.items():

    sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()

    sap_model.Results.Setup.SetCaseSelectedForOutput(val)

    [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = sap_model.Results.JointDispl(
        Frame3i, 0, 0, [], [], [], [], [], [], [], [], [], [], [])

    SapResult[val] = U1[0]

#%% CLOSE SAP2000 MODEL AND APPLICATION

save_model = False
my_sap_object = sap2000.closesap2000(my_sap_object, save_model)

#%% MANIPULATE DATA

# fill independent results

IndResult = OrderedDict([('LTYPE_DEAD', -0.02639), ('LTYPE_OTHER', 0.06296)])

# fill percent difference

PercentDiff = {}

for key, val in load_patterns.items():
    PercentDiff[val] = (SapResult[val] / IndResult[val]) - 1

# display results

for key, val in load_patterns.items():
    print()

    print(SapResult[val])

    print(IndResult[val])

    print(PercentDiff[val])
