import os
import sys
import comtypes.client
from collections import OrderedDict
import pandas as pd
# import numpy as np
# import matplotlib.pyplot as plt

#%% INITIALIZE COM CODE TO TIE INTO SAP2000 CSI OAPI

# set the following flag to True to attach to an existing instance of the program

# otherwise a new instance of the program will be started

AttachToInstance = False

# set the following flag to True to manually specify the path to SAP2000.exe

# this allows for a connection to a version of SAP2000 other than the latest installation

# otherwise the latest installed version of SAP2000 will be launched

SpecifyPath = False

# if the above flag is set to True, specify the path to SAP2000 below

ProgramPath = r"C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe"

# full path to the model

# set it to the desired path of your model

APIPath = r"C:\Users\Felipe\PycharmProjects\CoupledStructures\models"

if not os.path.exists(APIPath):

    try:

        os.makedirs(APIPath)

    except OSError:

        pass

#%% OPEN SAP2000 OR ATTACH TO EXISTING OPEN INSTANCE AND INSTANTIATE SAP2000 OBJECT

if AttachToInstance:

    # attach to a running instance of SAP2000

    try:

        # get the active SapObject

        mySapObject = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")

    except (OSError, comtypes.COMError):

        print("No running instance of the program found or failed to attach.")

        sys.exit(-1)

else:

    # create API helper object

    helper = comtypes.client.CreateObject('SAP2000v20.Helper')

    helper = helper.QueryInterface(comtypes.gen.SAP2000v20.cHelper)

    if SpecifyPath:

        try:

            # 'create an instance of the SAPObject from the specified path

            mySapObject = helper.CreateObject(ProgramPath)

        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program from " + ProgramPath)

            sys.exit(-1)

    else:

        try:

            # create an instance of the SAPObject from the latest installed SAP2000

            mySapObject = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

        except (OSError, comtypes.COMError):

            print("Cannot start a new instance of the program.")

            sys.exit(-1)

    # start SAP2000 application

    mySapObject.ApplicationStart()

#%% CREATE NEW SAP2000 MODEL AND INITIALIZE IT

# create SapModel object

SapModel = mySapObject.SapModel

# initialize model

SapModel.InitializeNewModel()

# create new blank model

ret = SapModel.File.NewBlank()

#%% DEFINE MODEL GEOMETRY, PROPERTIES, AND LOADING

# set model degrees of freedom to only allow forces and motion in the XZ plane

DOF = [True, False, True, False, True, False]

ret = SapModel.Analyze.SetActiveDOF(DOF)

# define material property

MATERIAL_STEEL = 1

ret = SapModel.PropMaterial.SetMaterial('STEEL', MATERIAL_STEEL)

# assign isotropic mechanical properties to material

ret = SapModel.PropMaterial.SetMPIsotropic('STEEL', 29000, 0.3, 6E-06)

# initialize section properties panda

modelPropFrame = pd.DataFrame(columns=['name', 'material', 'depth', 'width'])



ret = SapModel.PropFrame.SetRectangle('R1', 'STEEL', 12, 12)
ret = SapModel.PropFrame.SetRectangle('R2', 'STEEL', 60, 60)
ret = SapModel.PropFrame.SetRectangle('R3', 'STEEL', 12, 12)
ret = SapModel.PropFrame.SetRectangle('R4', 'STEEL', 60, 60)

# define frame section property modifiers

ModValue = [1, 1, 1, 1, 1, 1, 1, 1]

ret = SapModel.PropFrame.SetModifiers('R1', ModValue)

# switch to k-ft units

kip_ft_F = 4

ret = SapModel.SetPresentUnits(kip_ft_F)

# add frame object by coordinates

[Frame1, ret] = SapModel.FrameObj.AddByCoord(0, 0, 0, 0, 0, 20,  '', 'R1', 'Frame1', 'Global')
[Frame2, ret] = SapModel.FrameObj.AddByCoord(20, 0, 0, 20, 0, 20, '', 'R1', 'Frame2', 'Global')
[Frame3, ret] = SapModel.FrameObj.AddByCoord(0, 0, 20, 20, 0, 20, '', 'R2', 'Frame3', 'Global')
[Frame4, ret] = SapModel.FrameObj.AddByCoord(40, 0, 0, 40, 0, 20, '', 'R3', 'Frame4', 'Global')
[Frame5, ret] = SapModel.FrameObj.AddByCoord(60, 0, 0, 60, 0, 20, '', 'R3', 'Frame5', 'Global')
[Frame6, ret] = SapModel.FrameObj.AddByCoord(40, 0, 20, 60, 0, 20, '', 'R4', 'Frame6', 'Global')

# assign point object restraint at base

Restraint = [True, True, True, True, True, True]

[Frame1i, Frame1j, ret] = SapModel.FrameObj.GetPoints(Frame1, '', '')

ret = SapModel.PointObj.SetRestraint(Frame1i, Restraint)

[Frame2i, Frame1j, ret] = SapModel.FrameObj.GetPoints(Frame2, '', '')

ret = SapModel.PointObj.SetRestraint(Frame2i, Restraint)

[Frame4i, Frame4j, ret] = SapModel.FrameObj.GetPoints(Frame4, '', '')

ret = SapModel.PointObj.SetRestraint(Frame4i, Restraint)

[Frame5i, Frame5j, ret] = SapModel.FrameObj.GetPoints(Frame5, '', '')

ret = SapModel.PointObj.SetRestraint(Frame5i, Restraint)

# refresh view, update (initialize) zoom

ret = SapModel.View.RefreshView(0, False)

# add load patterns

LTYPE = OrderedDict([(1, 'D'), (5, 'EQ')])

ret = SapModel.LoadPatterns.Add(LTYPE[1], 1, 0, True)
ret = SapModel.LoadPatterns.Add(LTYPE[5], 5, 0, True)

# assign loading for load pattern 'D'

[Frame3i, Frame3j, ret] = SapModel.FrameObj.GetPoints(Frame3, '', '')
ret = SapModel.PointObj.SetLoadForce(Frame3i, LTYPE[1], [0, 0, -10, 0, 0, 0])
ret = SapModel.PointObj.SetLoadForce(Frame3j, LTYPE[1], [0, 0, -10, 0, 0, 0])

# assign loading for load pattern 'EQ'
[Frame3i, Frame3j, ret] = SapModel.FrameObj.GetPoints(Frame3, '', '')
ret = SapModel.FrameObj.SetLoadDistributed(Frame3, LTYPE[5], 1, 10, 0, 1, 1.8, 1.8)

# switch to k-in units

kip_in_F = 3

ret = SapModel.SetPresentUnits(kip_in_F)

#%% SAVE MODEL AND RUN IT

# save model

ModelPath = APIPath + os.sep + 'API_1-001.sdb'

ret = SapModel.File.Save(ModelPath)

# run model (this will create the analysis model)

ret = SapModel.Analyze.RunAnalysis()

#%% OBTAIN RESULTS FROM SAP2000 MODEL

# initialize for Sap2000 results

[Frame3i, Frame3j, ret] = SapModel.FrameObj.GetPoints(Frame3, '', '')

# get Sap2000 results for all load patterns:

SapResult = OrderedDict()

for key, val in LTYPE.items():

    ret = SapModel.Results.Setup.DeselectAllCasesAndCombosForOutput()

    ret = SapModel.Results.Setup.SetCaseSelectedForOutput(val)

    [NumberResults, Obj, Elm, ACase, StepType, StepNum, U1, U2, U3, R1, R2, R3, ret] = SapModel.Results.JointDispl(
        Frame3i, 0, 0, [], [], [], [], [], [], [], [], [], [], [])

    SapResult[val] = U1[0]

#%% CLOSE SAP2000 MODEL AND APPLICATION

# close Sap2000

ret = mySapObject.ApplicationExit(False)

SapModel = None

mySapObject = None

#%% MANIPULATE DATA

# fill independent results

IndResult = OrderedDict([('D', -0.02639), ('EQ', 0.06296)])

# fill percent difference

PercentDiff = {}

for key, val in LTYPE.items():
    PercentDiff[val] = (SapResult[val] / IndResult[val]) - 1

# display results

for key, val in LTYPE.items():
    print()

    print(SapResult[val])

    print(IndResult[val])

    print(PercentDiff[val])
