from zipfile import ZipFile
import pandas as pd
import numpy as np
import io

def read_zip(zip_path, zip_extract):
    zf = ZipFile(zip_path)
    return zf.read(zip_extract)

def read_xlxs(path, sheet_name=0):
    df = pd.read_excel(path, sheet_name=sheet_name)
    return df

def to_csv(df: pd.DataFrame, index=False, output_path='test.csv'):
    return df.to_csv(output_path, index=index, header= False)  

def sub_df(df, export_rows, exclude_columns):
    return df.iloc[export_rows].drop(axis=1,columns=df.columns[exclude_columns])

def find_first(df, str_to_find) -> int:
    for index, row in df.iterrows():
        if(row.isin([str_to_find]).any()):
            return index
    return 0

def find_footer(df, start_index=0):
    for index, row in df.iloc[start_index:].iterrows():
        if(row.isnull().any()):
            return index
    return -1        

def main(zip_path, xlsx):
    table = read_xlxs(io.BytesIO(read_zip(zip_path, xlsx)))
    first_index = find_first(table, 'Crimes Against Property')+1
    footer_index = find_footer(table, start_index=first_index)
    
    export_rows = list(range(first_index, footer_index))
    exclude_columns = [find_first(table, 'Total Victims1')+1]
    to_csv(sub_df(table, export_rows=export_rows, exclude_columns=exclude_columns))
    return

if __name__ == "__main__":
    main('victims.zip', 'Victims_Age_by_Offense_Category_2022.xlsx')
    print('Resultado generado como test.csv')