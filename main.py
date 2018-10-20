import os
import sap2000
from modelclasses import Model
import pandas as pd
from math import pi, sqrt
from numpy import arange

# %% OPEN SAP2000 AND READY PROGRAM

# check model directory
root_dir = os.getcwd()

model_path = root_dir + r'\models'

sap2000.check_model_path(model_path)

# create sap2000 object in memory
# in case user wants to attach to running instance, set attach_to_instance=True
# OAPI can automatically find most recent installation starting with SAP2000 v19
# For other installs, set specify_path=True
program_path = r'C:\Program Files\Computers and Structures\SAP2000 20\SAP2000.exe'
sap_obj = sap2000.attachtoapi(attach_to_instance=False, specify_path=False, program_path=program_path)

# open sap2000
# to show sap2000 GUI, set visible=True
model_obj = Model(sap2000.opensap2000(sap_obj, visible=False))

# open new model in units of lb_in_F

model_obj.new()


# set up output array

out_df_args = dict.fromkeys(['file_name', 'no_frames',
                             'max_u1_frm1', 'max_u1_frm2',
                             'm1', 'm2',
                             'k1', 'k2', 'kp',
                             'user_T', 'sap_T',
                             'frm1_bm_stiff', 'frm2_bm_stiff',
                             'frm1_col_stiff', 'frm2_col_stiff',
                             'frm1_bm_weight', 'frm2_bm_weight'])

out_df = pd.DataFrame(columns=list(out_df_args.keys()))
out_df.set_index('file_name', inplace=True)

run_flags = [1, 2]
k1_range = arange(10000, 102500, 2500).tolist()
kp_range = arange(1000, 10250, 250).tolist()

for k1_loop in k1_range:

    for kp_loop in kp_range:

        for flag in run_flags:
            # %% DEFINE MODEL GEOMETRY, PROPERTIES, AND LOADING

            # set file name for current run

            file_name = 'k1-{}_kp-{}_frm-{}'.format(k1_loop / 1000, kp_loop / 1000, flag)

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

            frm1_col_stiff = k1_loop / 2
            frm1_bm_stiff = 50 * frm1_col_stiff

            frm2_col_stiff = 50000
            frm2_bm_stiff = 50 * frm2_col_stiff

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

            ke[0] = kp_loop
            kp = ke[0]

            link_prop = {'name': 'Default', 'dof': dof, 'fixed': fixed,
                         'ke': ke, 'ce': ce, 'dj2': dj2, 'dj3': dj3,
                         'ke_coupled': False, 'ce_coupled': False,
                         'notes': '', 'guid': 'Default', 'w': 0,
                         'm': 0, 'R1': 0, 'R2': 0, 'R3': 0}

            model_obj.props.add_link_df(link_prop)

            model_obj.props.load_link_df()

            # switch to k-ft units
            model_obj.switch_units(units=sap2000.UNITS['lb_ft_F'])

            # generate frames given arguments set into the gen_frm_df() function and load frames into sap2000

            no_frames = flag
            no_stories = 1
            frm_width = 20
            frm1_bm_weight = 125000
            frm2_bm_weight = 125000
            frm_height = 20
            frm_spacing = 20

            model_obj.geometry.gen_frm(no_frames=no_frames, no_stories=no_stories, frm_width=frm_width,
                                       frm_height=frm_height, frm_spacing=frm_spacing,
                                       frm1_bm_weight=frm1_bm_weight, frm2_bm_weight=frm2_bm_weight)

            model_obj.geometry.load_frm_df()

            if flag == 2:
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

            model_obj.switch_units(units=sap2000.UNITS['lb_in_F'])

            model_obj.saveandrun(model_path=model_path, file_name=file_name)

            model_obj.refresh_view()

            # %% OBTAIN DATA
            class Results:

                def __init__(self):
                    [self.number_results, self.obj, self.elm, self.a_case, self.step_type, self.step_num,
                     self.u1, self.u2, self.u3, self.r1, self.r2, self.r3, self.ret] = [None] * 13

                    [self.mode_no, self.load_case, self.period, self.frequency,
                     self.circ_freq, self.eigen_value] = [None] * 6

                def new_joint_displ(self, sap_function):
                    [self.number_results, self.obj, self.elm, self.a_case, self.step_type, self.step_num,
                     self.u1, self.u2, self.u3, self.r1, self.r2, self.r3, self.ret] = sap_function

                def new_modal_period(self, sap_function):
                    [self.mode_no, self.load_case, self.step_type,
                     self.step_num, self.period, self.frequency, self.circ_freq,
                     self.eigen_value, self.ret] = sap_function

            res = Results()

            model_obj.sap_obj.Results.Setup.DeselectAllCasesAndCombosForOutput()
            model_obj.sap_obj.Results.Setup.SetCaseSelectedForOutput('TIME_HISTORY', True)
            model_obj.sap_obj.Results.Setup.SetOptionModalHist(1)

            res.new_joint_displ(
                    model_obj.sap_obj.Results.JointDispl('4', 0, 0, [], [], [], [], [], [], [], [], [], [], []))
            try:
                max_u1_frm1 = max((tuple(map(abs, res.u1))))

            except ValueError:
                print('There were issues obtaining max_u1_frm1 in file {}'.format(file_name))

            m1 = frm1_bm_weight / sap2000.GRAVITY / 12
            k1 = frm1_col_stiff * 2
            user_T = 2 * pi * sqrt(m1 / k1)

            max_u1_frm2 = None
            m2 = None
            k2 = None
            if flag == 2:

                res.new_joint_displ(
                        model_obj.sap_obj.Results.JointDispl('6', 0, 0, [], [], [], [], [], [], [], [], [], [], []))

                try:
                    max_u1_frm2 = max((tuple(map(abs, res.u1))))

                except ValueError:
                    print('There were issues obtaining max_u1_frm2 in file {}'.format(file_name))

                m2 = frm2_bm_weight / sap2000.GRAVITY / 12
                k2 = frm2_col_stiff * 2

                user_w = (1 / (2 * m1 ** 2 * m2 ** 2)) * (
                        k2 * m1 ** 2 * m2 + kp * m1 ** 2 * m2 + k1 * m1 * m2 ** 2 + kp * m1 * m2 ** 2 -
                        sqrt(((-k2) * m1 ** 2 * m2 - kp * m1 ** 2 * m2 - k1 * m1 * m2 ** 2 - kp * m1 * m2 ** 2) ** 2 -
                             4 * (k1 * k2 * m1 ** 3 * m2 ** 3 + k1 * kp * m1 ** 3 * m2 ** 3 + k2 * kp * m1 ** 3 * m2 ** 3)))

                user_T = 2*pi/sqrt(user_w)

            elif flag == 1:
                frm2_bm_stiff = None
                frm2_col_stiff = None
                frm2_bm_weight = None
                kp = None

            model_obj.sap_obj.Results.Setup.DeselectAllCasesAndCombosForOutput()
            model_obj.sap_obj.Results.Setup.SetCaseSelectedForOutput('MODAL', True)
            model_obj.sap_obj.Results.Setup.SetOptionModeShape(1, 1)

            res.new_modal_period(model_obj.sap_obj.Results.ModalPeriod(1, ['MODAL'], [], [], [], [], [], []))

            sap_T = min(res.period)

            out_df.loc[file_name] = [no_frames,
                                     max_u1_frm1, max_u1_frm2,
                                     m1, m2,
                                     k1, k2, kp,
                                     user_T, sap_T,
                                     frm1_bm_stiff, frm2_bm_stiff,
                                     frm1_col_stiff, frm2_col_stiff,
                                     frm1_bm_weight, frm2_bm_weight]

            print('Finished running {} ...'.format(file_name))

            model_obj.reset()


print('Debugging Placeholder')

# %% CLOSE SAP2000 MODEL AND APPLICATION

sap_obj = sap2000.closesap2000(sap_obj, save_model=False)

# %% MANIPULATE DATA

print('Script has terminated...')
