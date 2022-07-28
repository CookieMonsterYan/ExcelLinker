# -*- coding: utf-8 -*-
"""
对Excel sheet表进行关联操作

@author: Cookie Monster
"""

import pandas as pd
import os

def merge_one_sheet_with_sheets(file:str, index_col:str, 
                   files:list,
                   **kwds)->pd.DataFrame:
    '''
    

    Parameters
    ----------
    file : str
        DESCRIPTION.
    index_col : str
        DESCRIPTION.
    files : list
        DESCRIPTION.
    **kwds : TYPE
        DESCRIPTION.

    Returns
    -------
    main_df : TYPE
        DESCRIPTION.

    '''
    if 'sheet' in kwds:
        main_df = pd.read_excel(file, sheet_name=kwds['sheet'], dtype={index_col:str})
    else:
        main_df = pd.read_excel(file, dtype={index_col:str})
        
    main_file_folder = os.path.dirname(file)
    main_file_name = os.path.splitext(os.path.basename(file))[0]

    main_df[index_col] = main_df[index_col].apply(lambda x:str(x).strip())
        
    for file in files:
        file_name = os.path.splitext(os.path.basename(file))[0]
        df = pd.read_excel(file, dtype={index_col:str})
        df[index_col] = df[index_col].apply(lambda x:str(x).strip())
        in_df = main_df.merge(df, how="left", on=index_col)
        in_file_name = f"{main_file_name}_merged_{file_name}"
        in_df.to_excel(os.path.join(main_file_folder, in_file_name + ".xlsx"), index=False)
        main_file_name = in_file_name
        main_df = in_df

    return main_df


def link_sheets_with_one_sheet(file:str, index_col:str, 
                   files:list,
                   **kwds)->dict:
    '''
    

    Parameters
    ----------
    file : str
        DESCRIPTION.
    index_col : str
        DESCRIPTION.
    files : list
        DESCRIPTION.
    **kwds : TYPE
        DESCRIPTION.

    Returns
    -------
    dict
        DESCRIPTION.

    '''
    if 'sheet' in kwds:
        main_df = pd.read_excel(file, sheet_name=kwds['sheet'], dtype={index_col:str})
    else:
        main_df = pd.read_excel(file, dtype={index_col:str})
        
    main_file_folder = os.path.dirname(file)
    main_file_name = os.path.splitext(os.path.basename(file))[0]
    
    main_df[index_col] = main_df[index_col].apply(lambda x:str(x).strip())
    
    results = dict()
    
    for file in files:
        file_name = os.path.splitext(os.path.basename(file))[0]
        df = pd.read_excel(file, dtype={index_col:str})
        df[index_col] = df[index_col].apply(lambda x:str(x).strip())
        df_notin = df[~df[index_col].isin(main_df[index_col])]
        notin_file_name = f"{file_name}_notin_{main_file_name}"
        df_notin.to_excel(os.path.join(main_file_folder, notin_file_name + ".xlsx"), index=False)
        results[f"{file_name}_notin_{main_file_name}"] = df_notin
        df_in = df[df[index_col].isin(main_df[index_col])]
        df_in = df_in.merge(main_df, how="left", on=index_col)
        in_file_name = f"{file_name}_in_{main_file_name}"
        df_in.to_excel(os.path.join(main_file_folder, in_file_name + ".xlsx"), index=False)
        results[f"{file_name}_in_{main_file_name}"] = df_in
        
    return results
    
def find_cant_link_datas(file:str, index_col:str, 
                   files:list,
                   **kwds)->pd.DataFrame:
    '''
    

    Parameters
    ----------
    file : str
        DESCRIPTION.
    index_col : str
        DESCRIPTION.
    files : list
        DESCRIPTION.
    **kwds : TYPE
        DESCRIPTION.

    Returns
    -------
    main_df : TYPE
        DESCRIPTION.

    '''
    if 'sheet' in kwds:
        main_df = pd.read_excel(file, sheet_name=kwds['sheet'], dtype={index_col:str})
    else:
        main_df = pd.read_excel(file, dtype={index_col:str})
    
    main_file_folder = os.path.dirname(file)
    main_file_name = os.path.splitext(os.path.basename(file))[0]
    
    main_df[index_col] = main_df[index_col].apply(lambda x:str(x).strip())
    
    for file in files:
        file_name = os.path.splitext(os.path.basename(file))[0]
        df = pd.read_excel(file, dtype={index_col:str})
        df[index_col] = df[index_col].apply(lambda x:str(x).strip())
        notin_df = main_df[~main_df[index_col].isin(df[index_col])]
        notin_file_name = f"{main_file_name}_notin_{file_name}"
        notin_df.to_excel(os.path.join(main_file_folder, notin_file_name + ".xlsx"), index=False)
        in_df = main_df[main_df[index_col].isin(df[index_col])]
        in_df = in_df.merge(df, how="left", on=index_col)
        in_file_name = f"{main_file_name}_in_{file_name}"
        in_df.to_excel(os.path.join(main_file_folder, in_file_name + ".xlsx"), index=False)
        main_file_name = notin_file_name
        main_df = notin_df
    
    return main_df




