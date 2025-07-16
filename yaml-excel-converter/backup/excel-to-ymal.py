import yaml
import pandas as pd
from collections import OrderedDict
import re


class MySafeDumper(yaml.SafeDumper):
    def increase_indent(self, flow=False, indentless=False):
        return super(MySafeDumper, self).increase_indent(flow, False)


def represent_dictionary(self, data):
    return self.represent_mapping("tag:yaml.org,2002:map", data.items())


MySafeDumper.add_representer(dict, represent_dictionary)


def excel_to_yaml(excel_file, output_yaml_file, column_order):
    df = pd.read_excel(excel_file)
    data = {"tcp": {"nodes": []}}

    for index, row in df.iterrows():
        node_dict = {}
        for col in column_order:
            if not pd.isna(row[col]) and row[col] != ".nan":
                if isinstance(row[col], str) and row[col].startswith("[{"):
                    # Parse the string as a list of dictionaries
                    try:
                        nested_data = eval(row[col])
                        if isinstance(nested_data, list) and all(
                            isinstance(item, dict) for item in nested_data
                        ):
                            node_dict[col] = nested_data
                    except:
                        node_dict[col] = row[col].to_string()
                else:
                    node_dict[col] = row[col]
        if node_dict:  # Only add to list if it has any valid data
            data["tcp"]["nodes"].append(node_dict)

    # Convert any OrderedDicts to standard dictionaries for YAML serialization
    data = convert_odict_to_dict(data)

    # Generate the YAML string
    yaml_str = yaml.dump(
        data,
        Dumper=MySafeDumper,
        default_flow_style=False,
        allow_unicode=True,
        indent=2,
    )

    # Replace 'true' with 'True' and 'false' with 'False'
    yaml_str = yaml_str.replace("true", "True").replace("false", "False")
    yaml_str = re.sub(
        r"(\s*local_port:\s*)\d+\.\d*",
        lambda m: m.group(1) + str(int(float(m.group(0).split(":")[1].strip()))),
        yaml_str,
    )
    yaml_str = re.sub(r"(\s*passive:\s*)1\.0", r"\1True", yaml_str)
    yaml_str = re.sub(r"(\s*passive:\s*)0\.0", r"\1False", yaml_str)
    yaml_str = re.sub(r"(\s*(remote_host):\s*)(\S+)", r'\1"\3"', yaml_str)
    yaml_str = re.sub(
        r"(\s*description:\s*)(.*?)(\s*)$", r'\1"\2"\3', yaml_str, flags=re.MULTILINE
    )

    # Write the YAML string to the output file
    with open(output_yaml_file, "w", encoding="utf-8") as file:
        file.write(yaml_str)


def convert_odict_to_dict(data):
    """Recursively convert OrderedDict to standard dict."""
    if isinstance(data, OrderedDict):
        data = dict(data)
        for k, v in data.items():
            data[k] = convert_odict_to_dict(v)
    elif isinstance(data, list):
        data = [convert_odict_to_dict(item) for item in data]
    return data


excel_file = r"C:\xxxx\temp\output.xlsx"
output_yaml_file = r"C:\xxxx\temp\output.yaml"


# Example usage
column_order = [
    "node",
    "remote_host",
    "rsap",
    "description",
    # ....
]


excel_to_yaml(excel_file, output_yaml_file, column_order)
