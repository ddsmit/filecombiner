import pandas as pd


def combine_excel_files(folder_path):
    return pd.concat(
        [
            pd.read_excel(file_path)
            for file_path in folder_path.glob('*.xls*')
        ]
    )


def save_dataframe(df, file_path):
    df.to_excel(file_path, index=False)