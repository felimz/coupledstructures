import pandas as pd
import sap2000


class Model:
    def __init__(self, props=None, members=None, joints=None):
        self.props = props or Props()
        self.members = members or Members()
        self.joints = joints or Joints()


class Props:

    def __init__(self):

        self.mdl_dof_df = pd.DataFrame([dict([(1, True), (2, False), (3, True), (4, False), (5, True), (6, False)])])
        self.mat_props_df = []
        self.frm_props_list = []
        self.frm_props_df = []

    # Model properties methods
    def load_mdl_dof_df(self, sap_model, dof='2-D'):

        if dof == '3-D':
            self.mdl_dof_df.iloc[[0]] = [dict([(1, True), (2, True), (3, True), (4, True), (5, True), (6, True)])]

        sap_model.Analyze.SetActiveDOF(self.mdl_dof_df.values.tolist()[0])

    # Material properties methods
    def load_mat_props_df(self, sap_model):

        self.mat_props_df = pd.DataFrame([{'name': 'STEEL', 'material': sap2000.MATERIAL_TYPES['MATERIAL_STEEL'],
                                           'youngs': 29000, 'poisson': 0.3, 't_coeff': 6E-06}])

        for index, row in self.mat_props_df.iterrows():

            name, material, youngs, poisson, t_coeff = [row['name'], row['material'], row['youngs'],
                                                        row['poisson'], row['t_coeff']]

            sap_model.PropMaterial.SetMaterial(name, material)

            sap_model.PropMaterial.SetMPIsotropic(name, youngs, poisson, t_coeff)

    # Frame properties methods
    def load_frm_props_df(self, sap_model):

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

            self.frm_props_list.append({'name': name_prefix + str(m), 'material': 'STEEL', 'depth': depth,
                                        'width': width})

        self.frm_props_df = pd.DataFrame(self.frm_props_list)

        for index, row in self.frm_props_df.iterrows():

            name, material, depth, width = [row['name'], row['material'], row['depth'], row['width']]

            sap_model.PropFrame.SetRectangle(name, material, depth, width)


class Members:
    def __init__(self):
        print('Members Class')


class Joints:
    def __init__(self):
        print('Joints Class')
