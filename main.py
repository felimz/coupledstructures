import os
import sap2000
from modelclasses import Model
import pandas as pd

# %% OPEN SAP2000 AND READY PROGRAM

# check model directory
root_dir = os.getcwd()

model_path = root_dir + r'\models'

sap2000.check_model_path(model_path)

# create sap2000 object in memory
# in case user wants to attach to running instance, set attach_to_instance=True
# program finds most recent installation after SAP2000 v19. For other installs, set specify_path=True
program_path = r'C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe'
sap_obj = sap2000.attachtoapi(attach_to_instance=True, specify_path=False, program_path=program_path)

# open sap2000
# to show sap2000 GUI, set visible=True
model_obj = Model(sap2000.opensap2000(sap_obj, visible=True))

# open new model in units of kip_in_F
model_obj.new()

# %% DEFINE MODEL GEOMETRY, PROPERTIES, AND LOADING

# initialize model object

# set degrees of freedom
model_obj.props.set_mdl_dof_df(dof='2-D')

# load degrees of freedom for 2-D motion into sap2000
model_obj.props.load_mdl_dof_df()

# set material properties
mat_prop = {'material': 'STEEL', 'material_id': sap2000.MATERIAL_TYPES['MATERIAL_STEEL'],
            'youngs': 29000, 'poisson': 0.3, 't_coeff': 6E-06, 'weight': 0}

model_obj.props.add_mat_df(mat_prop)

# load dataframe of material properties into sap2000
model_obj.props.load_mat_df()

frm1_col_stiff = 25
frm1_bm_stiff = 200
frm2_col_stiff = 25
frm2_bm_stiff = 200

# generate frame properties
model_obj.props.gen_frm(frm1_col_stiff, frm1_bm_stiff, frm2_col_stiff, frm2_bm_stiff)

# load dataframe of frame properties into sap2000
model_obj.props.load_frm_df()

# add link property

dof = [False] * 6
fixed = [0] * 6
ke = [0] * 6
ce = [0] * 6
dj2 = 0
dj3 = 0

dof[0] = True

ke[0] = 5

link_prop = {'name': 'Default', 'dof': dof, 'fixed': fixed,
             'ke': ke, 'ce': ce, 'dj2': dj2, 'dj3': dj3,
             'ke_coupled': False, 'ce_coupled': False,
             'notes': '', 'guid': 'Default', 'w': 0,
             'm': 0, 'R1': 0, 'R2': 0, 'R3': 0}

model_obj.props.add_link_df(link_prop)

model_obj.props.load_link_df()

# switch to k-ft units
model_obj.switch_units(units=sap2000.UNITS['kip_ft_F'])

# generate frames given arguments set into the gen_frm_df() function and load frames into sap2000

no_frames = 2
no_stories = 1
frm1_bm_weight = 20
frm2_bm_weight = 20

model_obj.geometry.gen_frm(no_frames=no_frames, no_stories=no_stories, frm_width=20, frm_height=20, frm_spacing=20,
                           frm1_bm_weight=frm1_bm_weight, frm2_bm_weight=frm2_bm_weight)

model_obj.geometry.load_frm_df()

# add link

model_obj.geometry.new_link()

model_obj.geometry.load_link_df()

# add restraints

model_obj.geometry.set_restraints('fixed')

# refresh sap2000 view to show elements

model_obj.refresh_view()

# add time history excitation

model_obj.loads.set_time_history()

# %% SAVE MODEL AND RUN IT

model_obj.switch_units(units=sap2000.UNITS['kip_in_F'])

model_obj.saveandrun(model_path=model_path,
                     file_name='FRM{}_S-col1{}-col2{}_M-bm1{}-bm2{}'.format(no_frames,
                                                                            frm1_col_stiff,
                                                                            frm2_col_stiff,
                                                                            frm1_bm_weight,
                                                                            frm2_bm_weight))

# %% OBTAIN DATA
model_obj.sap_obj.Results.Setup.DeselectAllCasesAndCombosForOutput()
model_obj.sap_obj.Results.Setup.SetCaseSelectedForOutput('TIME_HISTORY', True)
model_obj.sap_obj.Results.Setup.SetOptionModalHist(1)

model_obj.refresh_view()


class Results:

    def __init__(self):

        [self.number_results, self.obj, self.elm, self.a_case, self.step_type, self.step_num,
         self.u1, self.u2, self.u3, self.r1, self.r2, self.r3, self.ret] = [None]*13

    def new_run(self, sap_function):
        [self.number_results, self.obj, self.elm, self.a_case, self.step_type, self.step_num,
         self.u1, self.u2, self.u3, self.r1, self.r2, self.r3, self.ret] = sap_function


ro = Results()

ro.new_run(model_obj.sap_obj.Results.JointDispl('4', 0, 0,
                                                [], [], [],
                                                [], [], [],
                                                [], [], [],
                                                [], []))



# %% CLOSE SAP2000 MODEL AND APPLICATION

sap_obj = sap2000.closesap2000(sap_obj, save_model=False)

# %% MANIPULATE DATA

print('debugging line')
