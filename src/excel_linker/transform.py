# -*- coding: utf-8 -*-
"""
对Excel中一些数据进行转换

@author: Cookie Monster
"""

import pandas as pd
import os

def transform_code(file:str, code_col:str, 
                   code_explain_file:str, explain_code_col:str,
                   **kwds)->pd.DataFrame:
    '''
    

    Parameters
    ----------
    file : str
        DESCRIPTION.
    code_col : str
        DESCRIPTION.
    code_explain_file : str
        DESCRIPTION.
    explain_code_col : str
        DESCRIPTION.
    **kwds : TYPE
        DESCRIPTION.

    Returns
    -------
    new_df : TYPE
        A new dataframe that code has been explained.

    '''
    if 'sheet' in kwds:
        main_df = pd.read_excel(file, sheet_name=kwds['sheet'], dtype={code_col:str})
    else:
        main_df = pd.read_excel(file, dtype={code_col:str})
    
    main_df[code_col] = main_df[code_col].apply(lambda x:str(x).strip())
    
    if 'code_start_index' in kwds and 'code_length' in kwds:
        code_start_index = kwds['code_start_index']
        code_length = kwds['code_length']
        code_col2 = f"{code_col}_{code_start_index}_{code_length}" 
        main_df[code_col2] = main_df[code_col].apply(lambda x:x[code_start_index:code_start_index+code_length])
        code_col = code_col2
        
    if 'explain_sheet' in kwds:
        explain_df = pd.read_excel(code_explain_file, sheet_name=kwds['explain_sheet'], dtype={explain_code_col:str})
    else:
        explain_df = pd.read_excel(code_explain_file, dtype={explain_code_col:str})
    
    explain_df[explain_code_col] = explain_df[explain_code_col].apply(lambda x:str(x).strip())
    
    
    new_df = main_df.merge(explain_df, how="left", left_on=code_col, right_on=explain_code_col)
    
    return new_df
