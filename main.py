import sap2000
from collections import OrderedDict
import pandas as pd

#%% OPEN SAP2000 AND READY PROGRAM

# check sap2000 installation and attach to sap2000 COM
api_path = r"C:\Users\Felipe\PycharmProjects\CoupledStructures\models"
sap2000.checkinstall(api_path)

# create sap2000 object in memory
attach_to_instance = False
specify_path = False
program_path = r"C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe"
my_sap_object = sap2000.attachtoapi(attach_to_instance, specify_path, program_path)

# create sap2000 model in memory
sap_model = sap2000.newmodel(my_sap_object)


#%% DEFINE MODEL GEOMETRY, PROPERTIES, AND LOADING

# set model degrees of freedom to only allow forces and motion in the XZ plane

DOF = [True, False, True, False, True, False]

sap_model.Analyze.SetActiveDOF(DOF)

# define material property

MATERIAL_STEEL = 1

sap_model.PropMaterial.SetMaterial('STEEL', MATERIAL_STEEL)

# assign isotropic mechanical properties to material

sap_model.PropMaterial.SetMPIsotropic('STEEL', 29000, 0.3, 6E-06)

# initialize section properties panda

modelPropFrame = pd.DataFrame(columns=['name', 'material', 'depth', 'width'])


sap_model.PropFrame.SetRectangle('R1', 'STEEL', 12, 12)
sap_model.PropFrame.SetRectangle('R2', 'STEEL', 60, 60)
sap_model.PropFrame.SetRectangle('R3', 'STEEL', 12, 12)
sap_model.PropFrame.SetRectangle('R4', 'STEEL', 60, 60)

# define frame section property modifiers

ModValue = [1, 1, 1, 1, 1, 1, 1, 1]

sap_model.PropFrame.SetModifiers('R1', ModValue)

# switch to k-ft units

kip_ft_F = 4

sap_model.SetPresentUnits(kip_ft_F)

# add frame object by coordinates

[Frame1, ret] = sap_model.FrameObj.AddByCoord(0, 0, 0, 0, 0, 20,  '', 'R1', 'Frame1', 'Global')
[Frame2, ret] = sap_model.FrameObj.AddByCoord(20, 0, 0, 20, 0, 20, '', 'R1', 'Frame2', 'Global')
[Frame3, ret] = sap_model.FrameObj.AddByCoord(0, 0, 20, 20, 0, 20, '', 'R2', 'Frame3', 'Global')
[Frame4, ret] = sap_model.FrameObj.AddByCoord(40, 0, 0, 40, 0, 20, '', 'R3', 'Frame4', 'Global')
[Frame5, ret] = sap_model.FrameObj.AddByCoord(60, 0, 0, 60, 0, 20, '', 'R3', 'Frame5', 'Global')
[Frame6, ret] = sap_model.FrameObj.AddByCoord(40, 0, 20, 60, 0, 20, '', 'R4', 'Frame6', 'Global')

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

LTYPE = OrderedDict([(1, 'D'), (5, 'EQ')])

sap_model.LoadPatterns.Add(LTYPE[1], 1, 0, True)
sap_model.LoadPatterns.Add(LTYPE[5], 5, 0, True)

# assign loading for load pattern 'D'

[Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')
sap_model.PointObj.SetLoadForce(Frame3i, LTYPE[1], [0, 0, -10, 0, 0, 0])
sap_model.PointObj.SetLoadForce(Frame3j, LTYPE[1], [0, 0, -10, 0, 0, 0])

# assign loading for load pattern 'EQ'
[Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')
sap_model.FrameObj.SetLoadDistributed(Frame3, LTYPE[5], 1, 10, 0, 1, 1.8, 1.8)

# switch to k-in units

kip_in_F = 3

sap_model.SetPresentUnits(kip_in_F)

#%% SAVE MODEL AND RUN IT

sap2000.saveandrunmodel(sap_model, api_path, file_name='API_1-001.sdb')

# #%% OBTAIN RESULTS FROM SAP2000 MODEL
#
# # initialize for Sap2000 results
#
# [Frame3i, Frame3j, ret] = sap_model.FrameObj.GetPoints(Frame3, '', '')
#
# # get Sap2000 results for all load patterns:
#
# SapResult = OrderedDict()
#
# for key, val in LTYPE.items():
#
#     sap_model.Results.Setup.DeselectAllCasesAndCombosForOutput()
#
#     sap_model.Results.Setup.SetCaseSelectedForOutput(val)
#
#     [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = sap_model.Results.JointDispl(
#         Frame3i, 0, 0, [], [], [], [], [], [], [], [], [], [], [])
#
#     SapResult[val] = U1[0]
#
# #%% CLOSE SAP2000 MODEL AND APPLICATION
#
# my_sap_object = sap2000.closemodel(my_sap_object)
#
# #%% MANIPULATE DATA
#
# # fill independent results
#
# IndResult = OrderedDict([('D', -0.02639), ('EQ', 0.06296)])
#
# # fill percent difference
#
# PercentDiff = {}
#
# for key, val in LTYPE.items():
#     PercentDiff[val] = (SapResult[val] / IndResult[val]) - 1
#
# # display results
#
# for key, val in LTYPE.items():
#     print()
#
#     print(SapResult[val])
#
#     print(IndResult[val])
#
#     print(PercentDiff[val])
