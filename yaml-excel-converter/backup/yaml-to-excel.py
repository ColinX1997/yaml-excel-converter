import yaml
import pandas as pd

SPS_config_path = r"C:\test.yaml"


# Function to read YAML file and return DataFrame
def yaml_to_df(yaml_file):
    with open(yaml_file, "r") as file:
        lines = file.readlines()[5:-1]  # Skip the first 5 lines and the last line
        content = "".join(lines)
        data = yaml.safe_load(content)
        nodes = data["nodes"]
        df = pd.json_normalize(nodes)
    return df


# Function to write DataFrame to Excel file
def df_to_excel(df, excel_file):
    df.to_excel(excel_file, index=False)


# Example usage
yaml_file = r"C:\test.yaml"
excel_file = r"C:\test.xlsx"

df = yaml_to_df(yaml_file)
df_to_excel(df, excel_file)
