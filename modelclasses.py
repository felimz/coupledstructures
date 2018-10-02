import os
import pandas as pd
import sap2000


class Model:

    def __init__(self, my_sap_obj):
        self.sap_obj = my_sap_obj
        self.props = self.Props(self.sap_obj)
        self.members = self.Members(self.sap_obj)
        self.loads = self.Loads(self.sap_obj)
        self.links = self.Links(self.sap_obj)
        self.tools = self.Tools(self.sap_obj)

    def new(self):
        # re-initialize sap model
        self.sap_obj.InitializeNewModel(sap2000.UNITS['kip_in_F'])

        # create new blank model
        self.sap_obj.File.NewBlank()

    def reset(self):
        # re-instance props, members, and joint classes and initialize new sap2000 model
        self.props = self.Props(self.sap_obj)
        self.members = self.Members(self.sap_obj)
        self.loads = self.Loads(self.sap_obj)
        self.links = self.Links(self.sap_obj)
        self.tools = self.Tools(self.sap_obj)
        self.new()

    def saveandrun(self, api_path, file_name='API_1-001.sdb'):
        # save model
        model_path = ''.join([api_path, os.sep, file_name])

        self.sap_obj.File.Save(model_path)

        # run model (this will create the analysis model)
        self.sap_obj.Analyze.RunAnalysis()

    def switch_units(self, units):
        self.sap_obj.SetPresentUnits(units)

    class Props:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            self.mdl_dof_df = pd.DataFrame(columns=[1, 2, 3, 4, 5, 6])
            self.mat_df = pd.DataFrame(columns=['material_id', 'material', 'youngs', 'poisson', 't_coeff'])
            self.frm_df = pd.DataFrame(columns=['name', 'material', 'depth', 'width'])
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

        def gen_frm_df(self, frm1_col_stiff, frm1_bm_stiff, frm2_col_stiff, frm2_bm_stiff):

            frm_dict_default = {'name': 'default', 'material': 'STEEL', 'depth': 12,
                                'width': 12, 'modifiers': [1, 1, 1, 1, 1, 1, 1, 1]}

            self.add_frm_df([frm_dict_default])

            no_frames = 2

            frm_list = []
            for i in range(no_frames):
                col_stiff = (frm1_col_stiff, frm2_col_stiff)
                bm_stiff = (frm1_bm_stiff, frm2_bm_stiff)
                m, k = (0, 0)
                for j in range(2):

                    if j % 2 == 0:
                        m += 1
                        name = 'F{}_RECT_COL{}'.format(i + 1, m)
                        depth = 12 * col_stiff[i]
                        width = 12 * col_stiff[i]
                        modifiers = [1, 1, 1, 1, 1, 1, 1, 1]
                    else:
                        k += 1
                        name = 'F{}_RECT_BM{}'.format(i + 1, k)
                        depth = 12 * bm_stiff[i]
                        width = 12 * bm_stiff[i]
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

            self.frm_df = pd.DataFrame([link_dict])

        def load_link_df(self):

            for index, row in self.frm_df.iterrows():
                self.sap_obj.PropLink.SetLinear(row['name'], row['dof'], row['fixed'],
                                                row['ke'], row['ce'], row['dj2'], row['dj3'],
                                                row['ke_coupled'], row['ce_coupled'],
                                                row['notes'], row['guid'])

    class Members:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            self.frm_df = pd.DataFrame(columns=['xi', 'yi', 'zi', 'xj', 'yj', 'zj',
                                                'name', 'prop_name', 'user_name',
                                                'csys', 'frm_i', 'frm_j'])

        def add_frm_df(self, xi=None, yi=None, zi=None, xj=None, yj=None, zj=None,
                       name='', prop_name=None, user_name=None, csys='Global'):

            self.frm_df = self.frm_df.append(pd.DataFrame([{'xi': xi, 'yi': yi, 'zi': zi,
                                                            'xj': xj, 'yj': yj, 'zj': zj,
                                                            'name': name, 'prop_name': prop_name,
                                                            'user_name': user_name, 'csys': csys}]),
                                             ignore_index=True, sort=False)

        def gen_frm_df(self, no_frames=2, no_stories=10, frm_width=20, frm_height=20, frm_spacing=20):

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
                         bm_or_col='col', left_or_right=1, props='default'):

                frm_df_args = dict.fromkeys(
                        ['xi', 'yi', 'zi', 'xj', 'yj', 'zj', 'prop_name', 'user_name'])

                frm_df_args['user_name'] = 'frm{}_st{}_{}{}'.format(frame_no, story_no, bm_or_col, left_or_right)

                frm_df_args.update({'xi': i_coords[0], 'yi': i_coords[0], 'zi': i_coords[0]})

                # calculate j_coord from given i_coords
                j_coords = ()
                if bm_or_col == 'col':

                    j_coords = (i_coords[0], i_coords[1], i_coords[2] + length)

                elif bm_or_col == 'bm':

                    j_coords = (i_coords[0] + length, i_coords[1], i_coords[2])

                frm_df_args.update({'xi': i_coords[0], 'yi': i_coords[1], 'zi': i_coords[2]})
                frm_df_args.update({'xj': j_coords[0], 'yj': j_coords[1], 'zj': j_coords[2]})

                frm_df_args['prop_name'] = props

                return frm_df_args

            for i in range(no_frames):

                if i == 0:

                    for j in range(no_stories):
                        self.add_frm_df(**new_memb(i_coords=(0, 0, j * frm_height),
                                                   length=frm_height, bm_or_col='col', frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1))
                        self.add_frm_df(**new_memb(i_coords=(frm_width, 0, j * frm_height),
                                                   length=frm_height, bm_or_col='col', frame_no=i + 1,
                                                   left_or_right=2, story_no=j + 1))
                        self.add_frm_df(**new_memb(i_coords=get_coords('frm{}_st{}_col{}'.format(i + 1, j + 1, 1),
                                                                       memb_end='j'), length=frm_height, bm_or_col='bm',
                                                   frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1))

                if no_frames == 2 and i == 1:

                    for j in range(no_stories):
                        self.add_frm_df(**new_memb(i_coords=(frm_width + frm_spacing, 0, j * frm_height),
                                                   length=frm_height, bm_or_col='col', frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1))
                        self.add_frm_df(**new_memb(i_coords=(2 * frm_width + frm_spacing, 0, j * frm_height),
                                                   length=frm_height, bm_or_col='col', frame_no=i + 1,
                                                   left_or_right=2, story_no=j + 1))
                        self.add_frm_df(**new_memb(i_coords=get_coords('frm{}_st{}_col{}'.format(i + 1, j + 1, 1),
                                                                       memb_end='j'), length=frm_height, bm_or_col='bm',
                                                   frame_no=i + 1,
                                                   left_or_right=1, story_no=j + 1))

        def load_frm_df(self):
            for index, row in self.frm_df.iterrows():
                self.sap_obj.FrameObj.AddByCoord(row['xi'], row['yi'], row['zi'],
                                                 row['xj'], row['yj'], row['zj'],
                                                 row['name'], row['prop_name'],
                                                 row['user_name'], row['csys'])

                frm_i, frm_j, run = self.sap_obj.FrameObj.GetPoints(row['user_name'], '', '')

                self.frm_df.set_index('user_name', inplace=True)

                self.frm_df.loc[row['user_name'], 'frm_i'] = frm_i
                self.frm_df.loc[row['user_name'], 'frm_j'] = frm_j

                self.frm_df.reset_index(inplace=True)

    class Links:
        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            self.link_df = pd.DataFrame(columns=['xi', 'yi', 'zi', 'xj', 'yj', 'zj',
                                                 'name', 'is_single_joint', 'prop_name', 'user_name',
                                                 'csys', 'link_i', 'link_j'])

        # Link methods
        def add_link_df(self, xi=None, yi=None, zi=None, xj=None, yj=None, zj=None,
                        name='', is_single_joint=False, prop_name=None, user_name=None, csys='Global'):
            self.link_df = self.link_df.append(pd.DataFrame([{'xi': xi, 'yi': yi, 'zi': zi,
                                                              'xj': xj, 'yj': yj, 'zj': zj,
                                                              'name': name, 'is_single_joint': is_single_joint,
                                                              'prop_name': prop_name, 'user_name': user_name,
                                                              'csys': csys}]),
                                               ignore_index=True, sort=False)

        def load_link_df(self):
            for index, row in self.link_df.iterrows():
                self.sap_obj.LinkObj.AddByCoord(row['xi'], row['yi'], row['zi'],
                                                row['xj'], row['yj'], row['zj'],
                                                row['name'], row['is_single_joint'],
                                                row['prop_name'], row['user_name'],
                                                row['csys'])

                link_i, link_j, run = self.sap_obj.LinkObj.GetPoints(row['user_name'], '', '')

                self.link_df.set_index('user_name', inplace=True)

                self.link_df.loc[row['user_name'], 'link_i'] = link_i
                self.link_df.loc[row['user_name'], 'link_j'] = link_j

                self.link_df.reset_index(inplace=True)

    class Loads:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj

    class Tools:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj

        def refresh_view(self):
            self.sap_obj.View.RefreshView(0, False)
