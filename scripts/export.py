import pandas as pd

def export_data_excel(file_path, sheet_name='Feuil1'):
    """export data from an Excel file into a pandas DataFrame."""
    
    df = pd.read_excel(file_path, sheet_name=sheet_name)
    print(f"✅ Successfully loaded data from {file_path}")
    return df

def export_data_csv(file_path, encoding='latin1', sep=';'):
    """export data from a CSV file into a pandas DataFrame."""
    
    df = pd.read_csv(file_path, encoding=encoding, sep=sep, header=None)
    print(f"✅ Successfully loaded data from {file_path}")
    return df