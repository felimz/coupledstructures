import os
import pandas as pd
import sap2000


class Model:

    def __init__(self, my_sap_obj):
        self.sap_obj = my_sap_obj
        self.props = self.Props(self.sap_obj)
        self.members = self.Members(self.sap_obj)
        self.joints = self.Joints(self.sap_obj)

    def new(self):
        # re-initialize sap model
        self.sap_obj.InitializeNewModel(sap2000.UNITS['kip_in_F'])

        # create new blank model
        self.sap_obj.File.NewBlank()

    def reset(self):
        # re-instance props, members, and joint classes and initialize new sap2000 model
        self.props = self.Props(self.sap_obj)
        self.members = self.Members(self.sap_obj)
        self.joints = self.Joints(self.sap_obj)
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
            self.mdl_dof_df = pd.DataFrame([dict([(1, True), (2, True), (3, True),
                                                  (4, True), (5, True), (6, True)])])
            self.mat_df = pd.DataFrame()
            self.frm_df = pd.DataFrame()

        # Model properties methods
        def set_mdl_dof_df(self, dof='2-D'):

            if dof == '3-D':
                self.mdl_dof_df = pd.DataFrame([dict([(1, True), (2, True), (3, True),
                                                      (4, True), (5, True), (6, True)])])
            elif dof == '2-D':
                self.mdl_dof_df = pd.DataFrame([dict([(1, True), (2, True), (3, True),
                                                      (4, True), (5, True), (6, True)])])
            else:
                self.mdl_dof_df = pd.DataFrame(dof)

        def load_mdl_dof_df(self):

            self.sap_obj.Analyze.SetActiveDOF(next(iter(self.mdl_dof_df.values.tolist())))

        # Material properties methods
        def add_mat_df(self, mat_dict):

            if self.mat_df.empty:
                self.mat_df = pd.DataFrame([mat_dict])
            else:
                self.mat_df.append([mat_dict])

        def load_mat_df(self):

            for index, row in self.mat_df.iterrows():

                _material, _material_id, _youngs, _poisson, _t_coeff = [row['material'],
                                                                        row['material_id'],
                                                                        row['youngs'],
                                                                        row['poisson'],
                                                                        row['t_coeff']]

                self.sap_obj.PropMaterial.SetMaterial(_material, _material_id)

                self.sap_obj.PropMaterial.SetMPIsotropic(_material, _youngs, _poisson, _t_coeff)

        def gen_frm_df(self, no_levels=1, col_stiff_1=1, bm_stiff_1=1,
                       col_stiff_2=1, bm_stiff_3=1, link_stiff=1, link_level=1):

            j = 1
            k = 1
            frm_list = []
            for i in range(4):

                if i % 2 == 0:
                    name = 'RECT_COL'
                    depth = 12
                    width = 12
                    modifiers = [1, 1, 1, 1, 1, 1, 1, 1]
                    m = j
                    j += 1
                else:
                    name = 'RECT_BM'
                    depth = 60
                    width = 60
                    m = k
                    k += 1
                    modifiers = [1, 1, 1, 1, 1, 1, 1, 1]

                frm_dict = {'name': name + str(m), 'material': 'STEEL', 'depth': depth,
                            'width': width, 'modifiers': modifiers}

                frm_list.append(frm_dict)

            self.add_frm_df(frm_list)

        # Frame properties methods
        def add_frm_df(self, frm_list):

            self.frm_df = pd.DataFrame(frm_list)

        def load_frm_df(self):

            for index, row in self.frm_df.iterrows():

                _name, _material, _depth, _width = [row['name'], row['material'],
                                                    row['depth'], row['width']]

                self.sap_obj.PropFrame.SetRectangle(_name, _material, _depth, _width)

        def set_frm_mod(self, name, modifiers):
            # define frame section property modifiers
            self.frm_df.set_index('name')
            self.frm_df[name] = modifiers

        def load_frm_mod_df(self, name):

            self.frm_df.set_index('name')
            _modifiers = self.frm_df[name]
            self.sap_obj.PropFrame.SetModifiers(name, _modifiers)

    class Members:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            self.frm_df = pd.DataFrame()
            print('Members Class')

        def load_frm_def(self):

            for index, row in self.frm_df.iterrows():

                _xi, _yi, _zi, _xj, _yj, _zj, _name, _prop_name, _user_name, _csys = [row['xi'], row['yi'], row['zi'],
                                                                                      row['xj'], row['yj'], row['zj'],
                                                                                      row['name'], row['prop_name'],
                                                                                      row['user_name'], row['csys']]

                self.sap_obj.FrameObj.AddByCoord(_xi, _yi, _zi, _xj, _yj, _zj,
                                                 _name, _prop_name, _user_name, _csys)

    class Joints:

        def __init__(self, sap_obj):
            self.sap_obj = sap_obj
            print('Joints Class')
