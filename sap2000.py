import os
import sys
import comtypes.client


#%% CHECK WHETHER SAP2000 IS INSTALLED, AND SET WORKING PATH
def checkinstall(api_path=r"C:\Users\Felipe\PycharmProjects\CoupledStructures\models"):

    if not os.path.exists(api_path):

        try:

            os.makedirs(api_path)

        except OSError:

            pass


#%% INITIALIZE COM CODE TO TIE INTO SAP2000 CSI OAPI AND OPEN SAP2000
# OR ATTACH TO EXISTING OPEN INSTANCE AND INSTANTIATE SAP2000 OBJECT
def attachtoapi(attach_to_instance=False, specify_path=False,
                program_path=r"C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe"):

    if attach_to_instance:

        # attach to a running instance of SAP2000

        try:

            # get the active SapObject

            my_sap_object = comtypes.client.GetActiveObject("CSI.SAP2000.API.SapObject")

        except (OSError, comtypes.COMError):

            print("No running instance of the program found or failed to attach.")

            sys.exit(-1)

    else:

        # create API helper object

        helper = comtypes.client.CreateObject('SAP2000v20.Helper')

        helper = helper.QueryInterface(comtypes.gen.SAP2000v20.cHelper)

        if specify_path:

            try:

                # 'create an instance of the SAPObject from the specified path

                my_sap_object = helper.CreateObject(program_path)

            except (OSError, comtypes.COMError):

                print("Cannot start a new instance of the program from " + program_path)

                sys.exit(-1)

        else:

            try:

                # create an instance of the SAPObject from the latest installed SAP2000

                my_sap_object = helper.CreateObjectProgID("CSI.SAP2000.API.SapObject")

            except (OSError, comtypes.COMError):

                print("Cannot start a new instance of the program.")

                sys.exit(-1)

        # start SAP2000 application

    return my_sap_object


#%% START SAP2000 AND CREATE NEW BLANK MODEL IN MEMORY
def newmodel(my_sap_object):

    # create sap_model object

    my_sap_object.ApplicationStart()

    sap_model = my_sap_object.SapModel

    # initialize model

    sap_model.InitializeNewModel()

    # create new blank model

    sap_model.File.NewBlank()

    return sap_model


#%% SAVE MODEL AND RUN IT
def saveandrunmodel(sap_model, api_path, file_name='API_1-001.sdb'):

    # save model
    
    model_path = api_path + os.sep + file_name
    
    sap_model.File.Save(model_path)
    
    # run model (this will create the analysis model)
    
    sap_model.Analyze.RunAnalysis()


#%% CLOSE SAP2000 MODEL AND APPLICATION
def closemodel(my_sap_object):
    # close Sap2000
    
    my_sap_object.ApplicationExit(False)

    my_sap_object = []

    return my_sap_object
