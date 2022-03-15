import os
import sys
import pandas as pd
import warnings

warnings.simplefilter("ignore")
path = input("Enter directory containing spreadsheets [Press Enter to continue]: ")
if not os.path.isdir(path):
    sys.exit('Please enter a valid path to a directory.')

output_file = os.path.join(path, 'merged.xlsx')

excel_files = []
for root, dirs, files in os.walk(path):
    for file in files:
        if (file.endswith(".xlsx") or file.endswith(".xls")) and \
                (not file.startswith("~$")) and \
                (not file.startswith("CollectorToolsSummaryReport")) and \
                (not file.startswith("merged")) and \
                (os.path.join(root, file) != output_file):
            excel_files.append(os.path.join(root, file))

sheets = {}
for file in excel_files:
    excel_file = pd.ExcelFile(file)
    sheets_in_file = excel_file.sheet_names

    for sheet in sheets_in_file:
        if not sheets.get(sheet):
            sheets[sheet] = []
        sheets[sheet].append(file)

with pd.ExcelWriter(output_file) as writer:
    sql_vms = pd.DataFrame()
    file_vms = pd.DataFrame()

    for key, value in sheets.items():
        df_output = pd.DataFrame()

        for file in value:
            excel_file = pd.ExcelFile(file)
            df = excel_file.parse(sheet_name=key)

            if 'VM Name' in df.columns:
                vm_type_list = []
                for index, row in df.iterrows():
                    if 'sql' in row.get('VM Name').lower():
                        vm_type = "Possible SQL DB VM"
                    elif 'db' in row.get('VM Name').lower():
                        vm_type = "Possible DB VM"
                    elif 'ora' in row.get('VM Name').lower():
                        vm_type = "Possible Oracle DB VM"
                    elif 'oracle' in row.get('VM Name').lower():
                        vm_type = "Possible Oracle DB VM"
                    elif 'fs' in row.get('VM Name').lower():
                        vm_type = "Possible File Server VM"
                    elif 'file' in row.get('VM Name').lower():
                        vm_type = "Possible File Server VM"
                    elif 'sharepoint' in row.get('VM Name').lower():
                        vm_type = "Possible Sharepoint VM"
                    elif 'vcsa' in row.get('VM Name').lower():
                        vm_type = "Possible VMware VM"
                    elif 'veeam' in row.get('VM Name').lower():
                        vm_type = "Possible Backup VM"
                    elif 'cisco' in row.get('VM Name').lower():
                        vm_type = "Possible Cisco VM"
                    elif 'docker' in row.get('VM Name').lower():
                        vm_type = "Possible Docker VM"
                    elif 'domaincontroller' in row.get('VM Name').lower():
                        vm_type = "Possible AD VM"
                    elif 'dc' in row.get('VM Name').lower():
                        vm_type = "Possible AD VM"
                    elif 'active directory' in row.get('VM Name').lower():
                        vm_type = "Possible AD VM"
                    elif 'fortinet' in row.get('VM Name').lower():
                        vm_type = "Possible Security VM"
                    elif 'secure' in row.get('VM Name').lower():
                        vm_type = "Possible Security VM"
                    elif 'bastion' in row.get('VM Name').lower():
                        vm_type = "Possible Security VM"
                    else:
                        vm_type = 'Unidentified VM'
                    vm_type_list.append(vm_type)

                df['VM Type'] = vm_type_list
            df_output = pd.concat([df_output, df], ignore_index=True)

        df_output.drop_duplicates(ignore_index=True, inplace=True)
        df_output.to_excel(writer, sheet_name=key, index=False)
