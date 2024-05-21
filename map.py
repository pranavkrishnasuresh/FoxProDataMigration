import pandas as pd

# Load the Excel sheets
# Change to local directory locations!!!
legacy_sheet = pd.read_excel('/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Legacy/All POs Legacy.xlsx')
conversion_sheet = pd.read_excel('/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Mapping File/All POs.xlsx')
cloud_sheet = pd.read_excel('/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Cloud/All POs Cloud.xlsx')

conversion_sheet['Cloud'] = conversion_sheet['Cloud'].str.strip()
conversion_sheet['Legacy'] = conversion_sheet['Legacy'].str.strip()
cloud_sheet.columns = cloud_sheet.columns.str.strip()

conversion_dict = pd.Series(conversion_sheet['Cloud'].values, index=conversion_sheet['Legacy']).to_dict()

cloud_data = pd.DataFrame(columns=cloud_sheet.columns)

print("Conversion Dictionary:")
print(conversion_dict)

transfer_log = []
for legacy_title in legacy_sheet.columns:
    legacy_title_clean = legacy_title.strip()
    if legacy_title_clean in conversion_dict:
        cloud_title = conversion_dict[legacy_title_clean]
        cloud_title_clean = cloud_title.strip()

        if cloud_title_clean in cloud_data.columns:
            print(f"Transferring data from Legacy '{legacy_title_clean}' to Cloud '{cloud_title_clean}'")
            cloud_data[cloud_title_clean] = legacy_sheet[legacy_title]
        else:
            transfer_log.append(f"Cloud title '{cloud_title_clean}' not found in cloud_data columns")
    else:
        transfer_log.append(f"Legacy title '{legacy_title_clean}' not found in conversion dictionary")

cloud_data_preview = cloud_data.head()

# Change output path of file!!!
output_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/updated_po_cloud_sheet.xlsx'

cloud_data.to_excel(output_path, index=False)

print(f"Data transfer complete. Check '{output_path}'")
print("Transfer Log:")
print(transfer_log)
print("Cloud Data Preview:")
print(cloud_data_preview)