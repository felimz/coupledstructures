import os
import sys
import comtypes.client


# %% CHECK WHETHER SAP2000 IS INSTALLED, AND SET WORKING PATH
def checkinstall(api_path=r'C:\Users\Felipe\PycharmProjects\CoupledStructures\models'):
    if not os.path.exists(api_path):

        try:

            os.makedirs(api_path)

        except OSError:

            pass


# %% INITIALIZE COM CODE TO TIE INTO SAP2000 CSI OAPI AND OPEN SAP2000
# OR ATTACH TO EXISTING OPEN INSTANCE AND INSTANTIATE SAP2000 OBJECT
def attachtoapi(attach_to_instance=False, specify_path=False,
                program_path=r'C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe'):
    if attach_to_instance:

        # attach to a running instance of SAP2000

        try:

            # get the active SapObject

            my_sap_object = comtypes.client.GetActiveObject('CSI.SAP2000.API.SapObject')

        except (OSError, comtypes.COMError):

            print('No running instance of the program found or failed to attach.')

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

                print('Cannot start a new instance of the program from ' + program_path)

                sys.exit(-1)

        else:

            try:

                # create an instance of the SAPObject from the latest installed SAP2000

                my_sap_object = helper.CreateObjectProgID('CSI.SAP2000.API.SapObject')

            except (OSError, comtypes.COMError):

                print('Cannot start a new instance of the program.')

                sys.exit(-1)

        # start SAP2000 application

    return my_sap_object


# %% START SAP2000 AND CREATE NEW BLANK MODEL IN MEMORY
def opensap2000(my_sap_object, visible=False):
    # create sap_model object

    my_sap_object.ApplicationStart(UNITS['kip_in_F'], visible)

    sap_model = my_sap_object.SapModel

    return sap_model


# %% CLOSE SAP2000 MODEL AND APPLICATION
def closesap2000(my_sap_object, save_model=False):
    my_sap_object.ApplicationExit(save_model)

    my_sap_object = []

    return my_sap_object


# %% THESE FUNCTIONS ARE MEANT TO INTERACT WITH THE PANDAS ARRAYS IN THE MODEL OBJECT

GRAVITY = 32.17405

DOF = {
    0: 'UX',
    1: 'UY',
    2: 'UZ',
    3: 'RX',
    4: 'RY',
    5: 'RZ'}

FRAME_TYPES = {
    'PortalFrame': 0,
    'ConcentricBraced': 1,
    'EccentricBraced': 2}

EFRAME_TYPES = {
    0: 'PortalFrame',
    1: 'ConcentricBraced',
    2: 'EccentricBraced'}

UNITS = {
    'lb_in_F': 1,
    'lb_ft_F': 2,
    'kip_in_F': 3,
    'kip_ft_F': 4,
    'kN_mm_C': 5,
    'kN_m_C': 6,
    'kgf_mm_C': 7,
    'kgf_m_C': 8,
    'N_mm_C': 9,
    'N_m_C': 10,
    'Ton_mm_C': 11,
    'Ton_m_C': 12,
    'kN_cm_C': 13,
    'kgf_cm_C': 14,
    'N_cm_C': 15,
    'Ton_cm_C': 16}

EUNITS = {
    1: 'lb_in_F',
    2: 'lb_ft_F',
    3: 'kip_in_F',
    4: 'kip_ft_F',
    5: 'kN_mm_C',
    6: 'kN_m_C',
    7: 'kgf_mm_C',
    8: 'kgf_m_C',
    9: 'N_mm_C',
    10: 'N_m_C',
    11: 'Ton_mm_C',
    12: 'Ton_m_C',
    13: 'kN_cm_C',
    14: 'kgf_cm_C',
    15: 'N_cm_C',
    16: 'Ton_cm_C'}

EITEM_TYPE = {
    'Object': 0,
    'Group': 1,
    'SelectedObjects': 2}

MATERIAL_TYPES = {
    'MATERIAL_STEEL': 1,
    'MATERIAL_CONCRETE': 2,
    'MATERIAL_NODESIGN': 3,
    'MATERIAL_ALUMINUM': 4,
    'MATERIAL_COLDFORMED': 5,
    'MATERIAL_REBAR': 6,
    'MATERIAL_TENDON': 7}

STEEL_SUBTYPES = {
    'MATERIAL_STEEL_SUBTYPE_ASTM_A36': 1,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A53GrB': 2,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A500GrB_Fy42': 3,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A500GrB_Fy46': 4,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A572Gr50': 5,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A913Gr50': 6,
    'MATERIAL_STEEL_SUBTYPE_ASTM_A992_Fy50': 7,
    'MATERIAL_STEEL_SUBTYPE_CHINESE_Q235': 8,
    'MATERIAL_STEEL_SUBTYPE_CHINESE_Q345': 9,
    'MATERIAL_STEEL_SUBTYPE_INDIAN_Fe250': 10,
    'MATERIAL_STEEL_SUBTYPE_INDIAN_Fe345': 11,
    'MATERIAL_STEEL_SUBTYPE_EN100252_S235': 12,
    'MATERIAL_STEEL_SUBTYPE_EN100252_S275': 13,
    'MATERIAL_STEEL_SUBTYPE_EN100252_S355': 14,
    'MATERIAL_STEEL_SUBTYPE_EN100252_S450': 15}

ESTEEL_SUBTYPES = {
    1: 'MATERIAL_STEEL_SUBTYPE_ASTM_A36',
    2: 'MATERIAL_STEEL_SUBTYPE_ASTM_A53GrB',
    3: 'MATERIAL_STEEL_SUBTYPE_ASTM_A500GrB_Fy42',
    4: 'MATERIAL_STEEL_SUBTYPE_ASTM_A500GrB_Fy46',
    5: 'MATERIAL_STEEL_SUBTYPE_ASTM_A572Gr50',
    6: 'MATERIAL_STEEL_SUBTYPE_ASTM_A913Gr50',
    7: 'MATERIAL_STEEL_SUBTYPE_ASTM_A992_Fy50',
    8: 'MATERIAL_STEEL_SUBTYPE_CHINESE_Q235',
    9: 'MATERIAL_STEEL_SUBTYPE_CHINESE_Q345',
    10: 'MATERIAL_STEEL_SUBTYPE_INDIAN_Fe250',
    11: 'MATERIAL_STEEL_SUBTYPE_INDIAN_Fe345',
    12: 'MATERIAL_STEEL_SUBTYPE_EN100252_S235',
    13: 'MATERIAL_STEEL_SUBTYPE_EN100252_S275',
    14: 'MATERIAL_STEEL_SUBTYPE_EN100252_S355',
    15: 'MATERIAL_STEEL_SUBTYPE_EN100252_S450'}

EMATERIAL_TYPES = {
    1: 'MATERIAL_STEEL',
    2: 'MATERIAL_CONCRETE',
    3: 'MATERIAL_NODESIGN',
    4: 'MATERIAL_ALUMINUM',
    5: 'MATERIAL_COLDFORMED',
    6: 'MATERIAL_REBAR',
    7: 'MATERIAL_TENDON'}

EOBJECT_TYPES = {
    1: 'points',
    2: 'frames',
    3: 'cables',
    4: 'tendons',
    5: 'areas',
    6: 'solids',
    7: 'links'}

OBJECT_TYPES = {
    'points': 1,
    'frames': 2,
    'cables': 3,
    'tendons': 4,
    'areas': 5,
    'solids': 6,
    'links': 7}

LOAD_PATTERN_TYPES = {
    'LTYPE_DEAD': 1,
    'LTYPE_SUPERDEAD': 2,
    'LTYPE_LIVE': 3,
    'LTYPE_REDUCELIVE': 4,
    'LTYPE_QUAKE': 5,
    'LTYPE_WIND': 6,
    'LTYPE_SNOW': 7,
    'LTYPE_OTHER': 8,
    'LTYPE_MOVE': 9,
    'LTYPE_TEMPERATURE': 10,
    'LTYPE_ROOFLIVE': 11,
    'LTYPE_NOTIONAL': 12,
    'LTYPE_PATTERNLIVE': 13,
    'LTYPE_WAVE': 14,
    'LTYPE_BRAKING': 15,
    'LTYPE_CENTRIFUGAL': 16,
    'LTYPE_FRICTION': 17,
    'LTYPE_ICE': 18,
    'LTYPE_WINDONLIVELOAD': 19,
    'LTYPE_HORIZONTALEARTHPRESSURE': 20,
    'LTYPE_VERTICALEARTHPRESSURE': 21,
    'LTYPE_EARTHSURCHARGE': 22,
    'LTYPE_DOWNDRAG': 23,
    'LTYPE_VEHICLECOLLISION': 24,
    'LTYPE_VESSELCOLLISION': 25,
    'LTYPE_TEMPERATUREGRADIENT': 26,
    'LTYPE_SETTLEMENT': 27,
    'LTYPE_SHRINKAGE': 28,
    'LTYPE_CREEP': 29,
    'LTYPE_WATERLOADPRESSURE': 30,
    'LTYPE_LIVELOADSURCHARGE': 31,
    'LTYPE_LOCKEDINFORCES': 32,
    'LTYPE_PEDESTRIANLL': 33,
    'LTYPE_PRESTRESS': 34,
    'LTYPE_HYPERSTATIC': 35,
    'LTYPE_BOUYANCY': 36,
    'LTYPE_STREAMFLOW': 37,
    'LTYPE_IMPACT': 38,
    'LTYPE_CONSTRUCTION': 39,
}
