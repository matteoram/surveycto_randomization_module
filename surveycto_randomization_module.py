#!/usr/bin/env python3
# -*- coding: utf-8 -*-

'''
project: surveycto randomization module
author: matteo ramina
created on: march 6, 2023
updated by: matteo ramina
updated on: september 14, 2023
'''

### modules
''' make sure to clone the github repository and run the command "pip3 install -r requirements.txt" 
 your terminal to install the required packages '''

import os
import math
import itertools as it
from pathlib import Path
import pandas as pd
import numpy as np
from pandas.io.formats import excel
import xlsxwriter
import openpyxl

### directory and output files
''' parameters to be changed manually '''

directory = Path('add_your_directory_here')
csv_name = 'surveycto_randomization_module.csv'
xlsx_name = 'surveycto_randomization_module.xlsx'
workbook_name = directory/xlsx_name

# functions

def verify_directory(dir_in):
    ''' verify that directory exists '''

    directory_correct = os.path.exists(dir_in)
    
    return directory_correct


def input_size():
    ''' input number of texts to randomize '''

    size_correct = False
    
    while not size_correct:
    
        size_out = input('Please enter the number of texts to randomize:')
    
        try:
    
            int(size_out)
            size_correct = True
            size_out = int(size_out)
    
        except ValueError:
    
            print('Please, enter an integer.')
    
    return size_out


def permutations(dir_in, size_in, csv_name):
    ''' create all permutations given 'size_in' and save them in 'dir_in'  '''
    
    # create permutations and store them in a Pandas dataframe
    list_perms = list(it.permutations(list(range(size_in))))
    df = pd.DataFrame(list_perms)
    
    # add 'v' in front of variables' names since they have to start with a letter
    df = df.add_prefix('v')
    
    # force index to start from 1 rather than 0
    df.index = np.arange(1, len(df)+1)
    
    # export file
    file_out = dir_in/csv_name
    df.to_csv(file_out, sep=',', index=True, index_label='permutations')
 
def texts(size_in):
    ''' enter texts to randomize and save them into a dictionary '''
   
    # create dictionary that will store the texts
    texts_out = {}
    print(f'Please enter the {size_in} texts to be randomized.\n')
   
    # enter as many texts as specified by 'size_in'  
    for i in range(size_in):
   
         text = input(f'Insert text {i+1}:')
         print(f'Text {i+1} is \"{text}\".\n')
         texts_out[i] = text
   
    return dict(texts_out)

def choices_tab(dict_in, size_in):
    ''' create the 'choices' tab to be copied in the surveycto questionnaire '''
    
    # create a Pandas dataframe with headings as in a stardard surveycto file's choices tab
    col_names = ['list_name', 'value', 'label']
    choices_out = pd.DataFrame(columns=col_names, index=range(size_in))
    
    # populate the dataframe given the texts in 'dict_in'
    choices_out['list_name'] = 'texts'
    choices_out['value'] = np.arange(1, choices_out.shape[0] + 1)
    choices_out['label'].update(pd.Series(dict_in))
    
    return choices_out

def survey_tab(name_in, size_in):
    ''' create the 'survey' tab to be copied in the surveycto questionnare '''
    
    # enter the field type associated with the text to randomize
    type_in = input('Enter the field type of the texts (for example \"integer\", \"select_one\", \"select_multiple\".\nIf \"select_one\" or \"select_multiple\" is entered, remember to input the list name too):')
    print(f'The field type entered is \"{type_in}\".\n')
    
    # initiate the worksheet and populate it
    workbook_out = xlsxwriter.Workbook(name_in)
    sheet1 = workbook_out.add_worksheet('survey')
    
    n = math.perm(size_in)
    
    sheet1.write('A1', 'type')
    sheet1.write('B1', 'name')
    sheet1.write('C1', 'label')
    sheet1.write('D1', 'relevance')
    sheet1.write('E1', 'calculation')
    sheet1.write('A2', 'select_one texts')
    sheet1.write('B2', 'texts')
    sheet1.write('C2', 'Field used to load the reference to the various texts')
    sheet1.write('D2', 'no')
    sheet1.write('A3', 'calculate')
    sheet1.write('B3', 'permutations_max')
    sheet1.write('C3', 'Maximum number of permutations')
    sheet1.write('E3', n)
    sheet1.write('A4', 'calculate')
    sheet1.write('B4', 'permutation_number')
    sheet1.write('C4', 'Random number generator')
    sheet1.write('E4', 'once(random())')
    sheet1.write('A5', 'calculate')
    sheet1.write('B5', 'permutation_selection')
    sheet1.write('C5', 'Selection of permutation based on random number generated')
    sheet1.write('E5', 'if(${permutation_number} = 1, ${permutation_max}, int(${permutation_number}*${permutation_max})+1)')

    j = 0

    for i in range(6, size_in*2+6, 2):

        j = j + 1

        type_code = 'calculate'
        type_label = 'calculate_here'
        name_code = 'text_' + str(j) + '_code'
        name_label = 'text_' + str(j) + '_label'
        label_code = 'Text ' + str(j) + ': code'
        label_label = 'Text ' + str(j) + ': label'
        calculation_code = 'pulldata(\"randomization\", \"v' + str(j) + '\", \"permutation\", ${permutation_selection})'
        calculation_label = 'jr:choice-name(${text_' + str(j) + '_code}, \"${texts}\")'

        cell_A_code = 'A' + str(i)
        cell_A_label = 'A' + str(i+1)
        cell_B_code = 'B' + str(i)
        cell_B_label = 'B' + str(i+1)
        cell_C_code = 'C' + str(i)
        cell_C_label = 'C' + str(i+1)
        cell_E_code = 'E' + str(i)
        cell_E_label = 'E' + str(i+1)

        sheet1.write(cell_A_code, type_code)
        sheet1.write(cell_A_label, type_label)
        sheet1.write(cell_B_code, name_code)
        sheet1.write(cell_B_label, name_label)
        sheet1.write(cell_C_code, label_code)
        sheet1.write(cell_C_label, label_label)
        sheet1.write(cell_E_code, calculation_code)
        sheet1.write(cell_E_label, calculation_label)

    row_group_start = size_in*2+6

    cell_A = 'A' + str(row_group_start)
    cell_B = 'B' + str(row_group_start)
    cell_C = 'C' + str(row_group_start)

    sheet1.write(cell_A, 'begin_group')
    sheet1.write(cell_B, 'randomization_group')
    sheet1.write(cell_C, 'Randomization module')

    cell_A = 'A' + str(row_group_start+1)
    cell_B = 'B' + str(row_group_start+1)
    cell_C = 'C' + str(row_group_start+1)

    sheet1.write(cell_A, 'note')
    sheet1.write(cell_B, 'randomization_note')
    sheet1.write(cell_C, 'Note of randomization module')

    j = 0

    for i in range(row_group_start+2, row_group_start+size_in+2):
        
        j = j + 1

        name = 'text_' + str(j)
        label = str(j) + '. ${text_' + str(j) + '_label}'

        cell_A = 'A' + str(i)
        cell_B = 'B' + str(i)
        cell_C = 'C' + str(i)

        sheet1.write(cell_A, type_in)
        sheet1.write(cell_B, name)
        sheet1.write(cell_C, label)

    row_group_end = row_group_start+size_in+2

    cell_A = 'A' + str(row_group_end)
    cell_B = 'B' + str(row_group_end)
    cell_C = 'C' + str(row_group_end)    

    sheet1.write(cell_A, 'end_group')
    sheet1.write(cell_B, 'randomization_group')
    sheet1.write(cell_C, 'Randomization module')
    
    # format changes
    sheet1.autofit()

    # save workbook
    workbook_out.close()

# run program

if __name__ == "__main__":
    
    directory_status = verify_directory(directory)

    if directory_status == True:

        size = input_size()

        permutations(directory, size, csv_name)

        list_texts = texts(size)

        choices = choices_tab(list_texts, size)

        survey_tab(workbook_name, size)

        # Combine 'survey' tab and 'choices' tab
        with pd.ExcelWriter(workbook_name, engine='openpyxl', mode='a') as writer:

            excel.ExcelFormatter.header_style = None
            choices.to_excel(writer, index=False, sheet_name='choices')

        print(f'Success! The files have been saved in \"{directory}\".')

    else:

        print('The directory does not exist, please enter a valid directory in row 28.\n')

# end