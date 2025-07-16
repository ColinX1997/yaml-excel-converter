import yaml
import pandas as pd
import os
import subprocess


class YamlToExcelConverter:
    def __init__(self, yaml_file, excel_file):
        self.yaml_file = yaml_file
        self.excel_file = excel_file

    def yaml_to_df(self):
        with open(self.yaml_file, "r") as file:
            lines = file.readlines()[5:-1]  # Skip the first 5 lines and the last line
            content = "".join(lines)
            data = yaml.safe_load(content)
            nodes = data["nodes"]
            df = pd.json_normalize(nodes)
        return df

    def df_to_excel(self, df):
        df.to_excel(self.excel_file, index=False)
        self.open_excel_file_after_saving()

    def open_excel_file_after_saving(self):
        # Use the appropriate command to open the Excel file based on the OS
        if os.name == "nt":
            os.startfile(self.excel_file)

    def convert_yaml_to_excel(self):
        df = self.yaml_to_df()
        self.df_to_excel(df)
