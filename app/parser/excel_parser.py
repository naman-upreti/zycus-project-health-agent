from pathlib import Path
import pandas as pd 


class ExcelParser:
    """
    Read an Excel workbook and Exposes its sheets 
    """
    
    def __init__(self,file_path:str):
        self.file_path = Path(file_path)
        
        
        if not self.file_path.exists():
            raise FileNotFoundError(f"File {self.file_path} does not exist")
        
        self.workbook = pd.ExcelFile(self.file_path)
        
    def get_sheet_names(self):
        """
        Returns the names of the sheets in the workbook
        """
        return self.workbook.sheet_names 
    
    def read_sheet(self,sheet_name:str):
        """
        Reads a sheet from the workbook and returns it as a DataFrame
        """
        if sheet_name not in self.workbook.sheet_names:
            raise ValueError(f"Sheet {sheet_name} does not exist in the workbook")
        
        return self.workbook.parse(sheet_name)
    
        
    def preview_workbook(self):
        """
        Display information about every sheet in the workbook.
        """

        print(f"\nWorkbook: {self.file_path.name}")
        print("=" * 60)

        for sheet in self.get_sheet_names():

            df = self.read_sheet(sheet)

            print(f"\nSheet: {sheet}")
            print(f"Rows: {df.shape[0]}")
            print(f"Columns: {df.shape[1]}")

            print("\nColumn Names:")

            for column in df.columns:
                print(f"  - {column}")

            print("-" * 60)