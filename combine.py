import pathlib
import pandas as pd

def get_path(message, check_exists=True):
    user_input = input(message)
    print(user_input)
    try:
        folder_path=pathlib.Path(user_input)
    except Exception:
        print("Not a valid path, please try again")
        return get_path(message, check_exists)

    if not check_exists:
        return folder_path

    if folder_path.exists():
        return folder_path
    else:
        print("Folder path does not exist, please try again")
        return get_path(message, check_exists)

def combine_excel_files(folder_path):
        return pd.concat(
            [
                pd.read_excel(file_path)
                for file_path in folder_path.glob('*.xls*')
            ]
        )

def save_dataframe(df, file_path):
    df.to_excel(file_path, index=False)

def combine_files_to_excel(folder_path=None, save_path=None):
    try:
        folder_path = folder_path or get_path("Enter the folder path to the excel files you want to combine ")
        save_path = save_path or get_path("Enter the folder path where you want to save the file ", check_exists=False)
        combined_files = combine_excel_files(folder_path=folder_path)
        save_dataframe(combined_files, save_path)
    except Exception as e:
        print(e)
        if 'y' in input('There was an error, do you want to try again? [y/n] ').lower():
            combine_files_to_excel()


if __name__=="__main__":
    combine_files_to_excel()

