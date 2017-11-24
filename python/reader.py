import pandas as pd

__author__ = 'venu'

if __name__ == '__main__':

    excel = pd.ExcelFile('/Users/venu/git/BeamScheduling/python/beamdatainput.xlsx', index_col=None, na_values=['NA'])
    beam_connectivity = excel.parse('ETABS_Input_Beam_Connectivity', skiprows=[0, 2], usecols=list(range(1, 5)))

    joint_coords = excel.parse('ETABS_Input_Joint_Coordinates', skiprows=[0, 2], usecols=list(range(1, 5)))
    frame_sections = excel.parse('ETABS_Input_Frame_Sections', skiprows=[0, 2], usecols=list(range(1, 6)))
    frame_assignments = excel.parse('ETABS_Input_Frame_Assignments', skiprows=[0, 2], usecols=list(range(1, 9)))

    print '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=beam_connectivity=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
    print beam_connectivity.index
    print beam_connectivity
    print '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=joint_coords=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
    print joint_coords.index
    print joint_coords
    print '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=frame_sections-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
    print frame_sections.index
    print frame_sections
    print '-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=frame_assignments-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-=-='
    print frame_assignments.index
    print frame_assignments