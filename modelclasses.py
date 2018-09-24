import os
import pandas as pd
import sap2000


class Model:
    def __init__(self, my_sap_object):

        self.sap_model = my_sap_object
        self.__resetvars()

    def __resetvars(self):

        self.props = Props()
        self.joints = Joints()
        self.members = Members()

    def new(self):

        # reset sub-classes
        self.__resetvars()

        # re-initialize sap model
        self.sap_model.InitializeNewModel(sap2000.UNITS['kip_in_F'])

        # create new blank model
        self.sap_model.File.NewBlank()

    def saveandrun(self, api_path, file_name='API_1-001.sdb'):
        # save model
        model_path = api_path + os.sep + file_name

        self.sap_model.File.Save(model_path)

        # run model (this will create the analysis model)
        self.sap_model.Analyze.RunAnalysis()


class Props(Model):

    def __init__(self):

        self.mdl_dof_df = pd.DataFrame([dict([(1, True), (2, False), (3, True), (4, False), (5, True), (6, False)])])
        self.mat_df = []
        self.frm_list = []
        self.frm_df = []

    # Model properties methods
    def load_mdl_dof_df(self, dof='2-D'):

        if dof == '3-D':
            self.mdl_dof_df.iloc[[0]] = [dict([(1, True), (2, True), (3, True), (4, True), (5, True), (6, True)])]

        self.sap_model.SetActiveDOF(self.mdl_dof_df.values.tolist()[0])

    # Material properties methods
    def load_mat_df(self):

        self.mat_df = pd.DataFrame([{'name': 'STEEL', 'material': sap2000.MATERIAL_TYPES['MATERIAL_STEEL'],
                                             'youngs': 29000, 'poisson': 0.3, 't_coeff': 6E-06}])

        for index, row in self.mat_df.iterrows():

            name, material, youngs, poisson, t_coeff = [row['name'], row['material'], row['youngs'],
                                                        row['poisson'], row['t_coeff']]

            self.sap_model.PropMaterial.SetMaterial(name, material)

            self.sap_model.PropMaterial.SetMPIsotropic(name, youngs, poisson, t_coeff)

    # Frame properties methods
    def load_frm_df(self):

        j = 1
        k = 1
        for i in range(4):

            if i % 2 == 0:
                name_prefix = 'RECT_COL'
                depth = 12
                width = 12
                m = j
                j += 1
            else:
                name_prefix = 'RECT_BM'
                depth = 60
                width = 60
                m = k
                k += 1

            self.frm_list.append({'name': name_prefix + str(m), 'material': 'STEEL', 'depth': depth,
                                  'width': width})

        self.frm_df = pd.DataFrame(self.frm_list)

        for index, row in self.frm_df.iterrows():

            name, material, depth, width = [row['name'], row['material'], row['depth'], row['width']]

            self.sap_model.PropFrame.SetRectangle(name, material, depth, width)


class Members(Model):

    def __init__(self):

        print('Members Class')


class Joints(Model):

    def __init__(self):

        print('Joints Class')

