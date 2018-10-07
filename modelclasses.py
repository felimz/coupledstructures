import os
import pandas as pd
import sap2000


class Model:

    def __init__(self, my_sap_obj):
        self.sap_obj = my_sap_obj
        self.props = self.Props(self.sap_obj)
        self.geometry = self.Geometry(self.sap_obj)
        self.loads = self.Loads(self.sap_obj)

    def new(self):
        # re-initialize sap model
        self.sap_obj.InitializeNewModel(sap2000.UNITS['kip_in_F'])

        # create new blank model
        self.sap_obj.File.NewBlank()

    def reset(self):
        # re-instance props, geometry, and joint classes and initialize new sap2000 model
        self.props = self.Props(self.sap_obj)
        self.geometry = self.Geometry(self.sap_obj)
        self.new()

    def saveandrun(self, api_path, file_name='API_1-001.sdb'):
        # save model
        model_path = ''.join([api_path, os.sep, file_name])

        self.sap_obj.File.Save(model_path)

        # run model (this will create the analysis model)
        self.sap_obj.Analyze.RunAnalysis()

    def switch_units(self, units):
        self.sap_obj.SetPresentUnits(units)

    def refresh_view(self):
        self.sap_obj.View.RefreshView(0, False)

    class Props:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            self.mdl_dof_df = pd.DataFrame(columns=[1, 2, 3, 4, 5, 6])
            self.mat_df = pd.DataFrame(columns=['material_id', 'material', 'youngs', 'poisson', 't_coeff', 'weight'])
            self.frm_df = pd.DataFrame(columns=['name', 'material', 'depth', 'width', 'mass'])
            self.link_df = pd.DataFrame(columns=['name', 'dof', 'fixed', 'ke', 'ce', 'dj2', 'dj3',
                                                 'ke_coupled', 'ce_coupled', 'notes', 'guid'])

        # Model properties methods
        def set_mdl_dof_df(self, dof='2-D'):

            if dof == '3-D':
                self.mdl_dof_df.loc[0] = dict([(1, True), (2, True), (3, True), (4, True), (5, True), (6, True)])

            elif dof == '2-D':
                self.mdl_dof_df.loc[0] = dict([(1, True), (2, False), (3, True), (4, False), (5, True), (6, False)])

            else:
                self.mdl_dof_df.loc[0] = dof

        def load_mdl_dof_df(self):

            self.sap_obj.Analyze.SetActiveDOF(next(iter(self.mdl_dof_df.values.tolist())))

        # Material properties methods
        def add_mat_df(self, mat_dict):

            self.mat_df = self.mat_df.append(pd.DataFrame([mat_dict]), ignore_index=True, sort=False)

        def load_mat_df(self):

            for index, row in self.mat_df.iterrows():
                self.sap_obj.PropMaterial.SetMaterial(row['material'], row['material_id'])

                self.sap_obj.PropMaterial.SetMPIsotropic(row['material'], row['youngs'],
                                                         row['poisson'], row['t_coeff'])

                self.sap_obj.PropMaterial.SetWeightAndMass(row['material'], 1, row['weight'])

        def gen_frm(self, frm1_col_stiff, frm1_bm_stiff, frm2_col_stiff, frm2_bm_stiff):

            frm_dict_default = {'name': 'Default', 'material': 'STEEL', 'depth': 12,
                                'width': 12, 'modifiers': [1, 1, 1, 1, 1, 1, 1, 1]}

            self.add_frm_df([frm_dict_default])

            no_frames = 2

            frm_list = []
            for i in range(no_frames):
                col_stiff = (frm1_col_stiff, frm2_col_stiff)
                bm_stiff = (frm1_bm_stiff, frm2_bm_stiff)
                m, k = (0, 0)
                for j in range(2):

                    # for columns
                    if j % 2 == 0:
                        m += 1
                        name = 'F{}_RECT_COL{}'.format(i + 1, m)
                        depth = (col_stiff[i] * 476.6896552) ** (1 / 4)
                        width = (col_stiff[i] * 476.6896552) ** (1 / 4)
                        modifiers = [1, 1, 1, 1, 1, 1, 1, 1]
                    # for beams
                    else:
                        k += 1
                        name = 'F{}_RECT_BM{}'.format(i + 1, k)
                        depth = (bm_stiff[i] * 476.6896552) ** (1 / 4)
                        width = (bm_stiff[i] * 476.6896552) ** (1 / 4)
                        modifiers = [1, 1, 1, 1, 1, 1, 1, 1]

                    frm_dict = {'name': name, 'material': 'STEEL', 'depth': depth,
                                'width': width, 'modifiers': modifiers}

                    frm_list.append(frm_dict)

            self.add_frm_df(frm_list)

        # Frame properties methods
        def add_frm_df(self, frm_list):

            self.frm_df = self.frm_df.append(pd.DataFrame(frm_list), ignore_index=True, sort=False)

        def load_frm_df(self):

            for index, row in self.frm_df.iterrows():
                self.sap_obj.PropFrame.SetRectangle(row['name'], row['material'],
                                                    row['depth'], row['width'])

        def set_frm_mod(self, name, modifiers):
            # define frame section property modifiers
            self.frm_df.set_index('name')
            self.frm_df[name] = modifiers

        def load_frm_mod_df(self, name):

            self.frm_df.set_index('name')
            _modifiers = self.frm_df[name]
            self.sap_obj.PropFrame.SetModifiers(name, _modifiers)

        # Link properties methods
        def add_link_df(self, link_dict):

            self.link_df = self.link_df.append(pd.DataFrame([link_dict]), ignore_index=True, sort=False)

        def load_link_df(self):

            for index, row in self.link_df.iterrows():
                self.sap_obj.PropLink.SetLinear(row['name'], row['dof'], row['fixed'],
                                                row['ke'], row['ce'], row['dj2'], row['dj3'],
                                                row['ke_coupled'], row['ce_coupled'],
                                                row['notes'], row['guid'])

                self.sap_obj.PropLink.SetWeightAndMass(row['name'], row['w'], row['m'],
                                                       row['R1'], row['R2'], row['R3'])

    class Geometry:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj

            self.frm_df = pd.DataFrame(columns=['xi', 'yi', 'zi', 'xj', 'yj', 'zj',
                                                'name', 'is_single_joint', 'prop_name', 'user_name',
                                                'csys', 'frm_i', 'frm_j', 'frm_type', 'frame_no',
                                                'story_no', 'i_restraint', 'j_restraint', 'mass'])

        def add_frm_df(self, xi=None, yi=None, zi=None, xj=None, yj=None, zj=None,
                       name='', prop_name=None, user_name=None, csys='Global',
                       frm_type=None, frame_no=None, story_no=None, mass=None):

            self.frm_df = self.frm_df.append(pd.DataFrame([{'xi': xi, 'yi': yi, 'zi': zi,
                                                            'xj': xj, 'yj': yj, 'zj': zj,
                                                            'name': name, 'prop_name': prop_name,
                                                            'user_name': user_name, 'csys': csys,
                                                            'frm_type': frm_type, 'frame_no': frame_no,
                                                            'story_no': story_no, 'mass': mass}]),
                                             ignore_index=True, sort=False)

        def gen_frm(self, no_frames=2, no_stories=10, frm_width=20, frm_height=20, frm_spacing=20,
                    frm1_bm_weight=20, frm2_bm_weight=20):

            def get_coords(frm_user_name=None, memb_end='i'):
                x, y, z = (0, 0, 0)
                if frm_user_name is None:
                    x, y, z = (0, 0, 0)
                if memb_end == 'i':
                    self.frm_df.set_index('user_name', inplace=True)
                    x, y, z = (self.frm_df.loc[frm_user_name, 'xi'],
                               self.frm_df.loc[frm_user_name, 'yi'],
                               self.frm_df.loc[frm_user_name, 'zi'])
                    self.frm_df.reset_index(inplace=True)
                elif memb_end == 'j':
                    self.frm_df.set_index('user_name', inplace=True)
                    x, y, z = (self.frm_df.loc[frm_user_name, 'xj'],
                               self.frm_df.loc[frm_user_name, 'yj'],
                               self.frm_df.loc[frm_user_name, 'zj'])
                    self.frm_df.reset_index(inplace=True)
                return x, y, z

            def new_memb(i_coords=(0, 0, 0), length=1, frame_no=1, story_no=1,
                         frm_type='col', left_or_right=1, props='Default', weight=0):

                frm_df_args = dict.fromkeys(
                        ['xi', 'yi', 'zi', 'xj', 'yj', 'zj',
                         'prop_name', 'user_name',
                         'frm_type', 'frame_no', 'story_no'])

                frm_df_args.update({'xi': i_coords[0], 'yi': i_coords[0], 'zi': i_coords[0]})

                # calculate j_coord from given i_coords
                j_coords = ()
                if frm_type == 'col':

                    j_coords = (i_coords[0], i_coords[1], i_coords[2] + length)

                elif frm_type == 'bm':

                    j_coords = (i_coords[0] + length, i_coords[1], i_coords[2])

                    frm_df_args['mass'] = weight / sap2000.GRAVITY / length

                frm_df_args.update({'xi': i_coords[0], 'yi': i_coords[1], 'zi': i_coords[2]})
                frm_df_args.update({'xj': j_coords[0], 'yj': j_coords[1], 'zj': j_coords[2]})

                frm_df_args['prop_name'] = props

                frm_df_args['user_name'] = 'frm{}_st{}_{}{}'.format(frame_no, story_no, frm_type, left_or_right)

                frm_df_args['frm_type'] = frm_type

                frm_df_args['frame_no'] = frame_no

                frm_df_args['story_no'] = story_no

                return frm_df_args

            for i in range(no_frames):

                if i == 0:

                    for j in range(no_stories):
                        self.add_frm_df(**new_memb(i_coords=(0, 0, j * frm_height),
                                                   length=frm_height, frm_type='col', frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1, props='F{}_RECT_COL1'.format(i + 1)
                                                   ))
                        self.add_frm_df(**new_memb(i_coords=(frm_width, 0, j * frm_height),
                                                   length=frm_height, frm_type='col', frame_no=i + 1,
                                                   left_or_right=2, story_no=j + 1, props='F{}_RECT_COL1'.format(i + 1)
                                                   ))
                        self.add_frm_df(**new_memb(i_coords=get_coords('frm{}_st{}_col{}'.format(i + 1, j + 1, 1),
                                                                       memb_end='j'), length=frm_height, frm_type='bm',
                                                   frame_no=i + 1, left_or_right=1, props='F{}_RECT_BM1'.format(i + 1),
                                                   story_no=j + 1, weight=frm1_bm_weight))

                if no_frames == 2 and i == 1:

                    for j in range(no_stories):
                        self.add_frm_df(**new_memb(i_coords=(frm_width + frm_spacing, 0, j * frm_height),
                                                   length=frm_height, frm_type='col', frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1,
                                                   props='F{}_RECT_COL1'.format(i + 1)))
                        self.add_frm_df(**new_memb(i_coords=(2 * frm_width + frm_spacing, 0, j * frm_height),
                                                   length=frm_height, frm_type='col', frame_no=i + 1,
                                                   left_or_right=2, story_no=j + 1,
                                                   props='F{}_RECT_COL1'.format(i + 1)))
                        self.add_frm_df(**new_memb(i_coords=get_coords('frm{}_st{}_col{}'.format(i + 1, j + 1, 1),
                                                                       memb_end='j'), length=frm_height, frm_type='bm',
                                                   frame_no=i + 1, left_or_right=1, props='F{}_RECT_BM1'.format(i + 1),
                                                   story_no=j + 1, weight=frm2_bm_weight))

        def load_frm_df(self):

            previous_units = self.sap_obj.GetPresentUnits()

            self.sap_obj.SetPresentUnits(sap2000.UNITS['kip_ft_F'])

            for index, row in self.frm_df.loc[self.frm_df['frm_type'] != 'link'].iterrows():
                self.sap_obj.FrameObj.AddByCoord(row['xi'], row['yi'], row['zi'],
                                                 row['xj'], row['yj'], row['zj'],
                                                 row['name'], row['prop_name'],
                                                 row['user_name'], row['csys'])

                frm_i, frm_j, run = self.sap_obj.FrameObj.GetPoints(row['user_name'], '', '')

                self.frm_df.set_index('user_name', inplace=True)

                self.frm_df.loc[row['user_name'], 'frm_i'] = frm_i
                self.frm_df.loc[row['user_name'], 'frm_j'] = frm_j

                self.frm_df.reset_index(inplace=True)

                # Add mass to frame object

                if row['frm_type'] == 'bm':
                    self.sap_obj.FrameObj.SetMass(row['user_name'], row['mass'], True, 0)

            self.sap_obj.SetPresentUnits(previous_units)

        # Link methods
        def add_link_df(self, frm_i=None, frm_j=None, xi=None, yi=None, zi=None, xj=None, yj=None, zj=None,
                        name='', is_single_joint=False, prop_name=None, user_name=None,
                        frm_type=None, story_no=None, csys='Global'):

            self.frm_df = self.frm_df.append(
                    pd.DataFrame([{'frm_i': frm_i, 'frm_j': frm_j, 'xi': xi, 'yi': yi, 'zi': zi,
                                   'xj': xj, 'yj': yj, 'zj': zj,
                                   'name': name, 'is_single_joint': is_single_joint,
                                   'prop_name': prop_name, 'user_name': user_name,
                                   'frm_type': frm_type, 'story_no': story_no,
                                   'csys': csys}]),
                    ignore_index=True, sort=False)

        def new_link(self, story_no=0, props='Default'):

            def get_coords(frm_user_name, memb_end='i'):

                x, y, z, frm_i_or_j = [None, None, None, None]

                if memb_end == 'i':
                    self.frm_df.set_index('user_name', inplace=True)
                    frm_i_or_j, x, y, z = (self.frm_df.loc[frm_user_name, 'frm_i'],
                                           self.frm_df.loc[frm_user_name, 'xi'],
                                           self.frm_df.loc[frm_user_name, 'yi'],
                                           self.frm_df.loc[frm_user_name, 'zi'])
                    self.frm_df.reset_index(inplace=True)
                elif memb_end == 'j':
                    self.frm_df.set_index('user_name', inplace=True)
                    frm_i_or_j, x, y, z = (self.frm_df.loc[frm_user_name, 'frm_j'],
                                           self.frm_df.loc[frm_user_name, 'xj'],
                                           self.frm_df.loc[frm_user_name, 'yj'],
                                           self.frm_df.loc[frm_user_name, 'zj'])
                    self.frm_df.reset_index(inplace=True)

                return x, y, z, frm_i_or_j

            story = None

            if story_no == 0:

                story = self.frm_df[['frame_no', 'story_no', 'zi', 'frm_type']].sort_values('zi', ascending=False).loc[
                    (self.frm_df['frame_no'] == 1) & (self.frm_df['frm_type'] == 'bm'), 'story_no'].iloc[0]

            elif story_no > 0:

                story = self.frm_df[['frame_no', 'story_no', 'zi', 'frm_type']].sort_values('zi', ascending=True).loc[
                    (self.frm_df['frame_no'] == 1) & (self.frm_df['frm_type'] == 'bm'), 'story_no'].iloc[story_no]

            frm_user_name_1 = 'frm{}_st{}_bm{}'.format(1, story, 1)  # this obtains the name of the top story on frm1
            frm_user_name_2 = 'frm{}_st{}_bm{}'.format(2, story, 1)  # this obtains the name of the top story on frm2

            xi, yi, zi, frm_i = get_coords(frm_user_name_1, memb_end='j')
            xj, yj, zj, frm_j = get_coords(frm_user_name_2, memb_end='i')

            self.add_link_df(frm_i=frm_i, frm_j=frm_j, xi=xi, yi=yi, zi=zi, xj=xj, yj=yi, zj=zj,
                             name='', is_single_joint=False, prop_name=props,
                             user_name='Default', csys='Global', frm_type='link', story_no=story)

        def load_link_df(self):

            previous_units = self.sap_obj.GetPresentUnits()

            self.sap_obj.SetPresentUnits(sap2000.UNITS['kip_ft_F'])

            for index, row in self.frm_df.loc[self.frm_df['frm_type'] == 'link'].iterrows():
                self.sap_obj.LinkObj.AddByPoint(row['frm_i'], row['frm_j'], row['name'],
                                                row['is_single_joint'], row['prop_name'], row['user_name'])

            self.sap_obj.SetPresentUnits(previous_units)

        def add_restraints_df(self, name=None, value=None):

            self.frm_df['i_restraint'].loc[self.frm_df['frm_i'] == name] = [value]

        def set_restraints(self, restraint_type='pinned'):

            restraints = self.frm_df.loc[(self.frm_df['frm_type'] == 'col') & (self.frm_df['story_no'] == 1), 'frm_i']

            value = None

            if restraint_type == 'pinned':
                value = [True, True, True, False, False, False]

            elif restraint_type == 'fixed':
                value = [True, True, True, True, True, True]

            for joint_name in restraints:
                self.add_restraints_df(joint_name, value)
                self.sap_obj.PointObj.SetRestraint(joint_name, value)

    class Loads:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj

        def new_time_history_function(self, time_history_file):
            pass

        def load_time_history(self, name, file_name, headlines, pre_chars,
                                                                   points_per_line, value_type, free_format, number_fixed, dt)

            self.sap_obj.Func.FuncTH.SetFromFile_1(name, file_name, headlines, pre_chars, points_per_line, value_type,
                                               free_format, number_fixed, dt)
