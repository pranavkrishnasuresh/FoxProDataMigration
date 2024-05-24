import ollama
import httpx
import pandas as pd

def load_and_prepare_sheets():
    # Load the Excel sheets
    legacy_sheet_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Legacy/Vendor Legacy.xlsx'
    conversion_sheet_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Mapping File/Vendor.xlsx'
    cloud_sheet_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Cloud/Vendor Cloud.xlsx'
    
    legacy_sheet = pd.read_excel(legacy_sheet_path)
    conversion_sheet = pd.read_excel(conversion_sheet_path)
    cloud_sheet = pd.read_excel(cloud_sheet_path)
    
    # Strip whitespace from columns in conversion_sheet
    conversion_sheet['Cloud'] = conversion_sheet['Cloud'].str.strip()
    conversion_sheet['Legacy'] = conversion_sheet['Legacy'].str.strip()
    cloud_sheet.columns = cloud_sheet.columns.str.strip()

    # Convert relevant columns in conversion_sheet to string type
    conversion_sheet['Cloud'] = conversion_sheet['Cloud'].astype(str)
    conversion_sheet['Legacy'] = conversion_sheet['Legacy'].astype(str)
    conversion_sheet = conversion_sheet.astype(str)

    # Prepare the data to be passed to the API
    conversion_data = conversion_sheet.apply(lambda row: f"{row['Cloud']}: {row['Legacy']}", axis=1).tolist()
    conversion_string = ", ".join(conversion_data)
    
    return conversion_string, conversion_sheet, legacy_sheet_path, cloud_sheet_path

def call_ollama_api(conversion_string, cloud_list, legacy_list):
    try:
        conversionMap = ollama.chat(model='llama3', messages=[
            {
                "role": "user", 
                "content": f'Im going to give you a conversion sheet which shows the mapping between the legacy and cloud columns. I want you to take this information and output the data in the following format. Cloud: Legacy, Cloud: Legacy. Don\'t include any data or introduction except for the final answer, don\'t provide any explanation and don\'t actually include Cloud: Legacy in your answer. Here is the data: {conversion_string}',
            },
        ])
        conversion_map_content = conversionMap['message']['content']
    except httpx.RequestError as e:
        print(f"An HTTP request error occurred: {e}")
        return None
    except httpx.HTTPStatusError as e:
        print(f"An HTTP status error occurred: {e.response.status_code} - {e.response.text}")
        return None
    except Exception as e:
        print(f"An unexpected error occurred: {e}")
        return None

    try:
        finalOutput = ollama.chat(model='llama3', messages=[
            {
                "role": "user", 
                "content": f'I\'m going to give you conversion equivalencies, most of these should be correct but I\'m not sure: {conversion_map_content}. However, here is the list of cloud titles: {cloud_list} and here is the list of legacy titles: {legacy_list}. If you think any of the conversion equivalencies are incorrect due to excessive similarities between titles in the legacy and cloud lists, please make these changes and output the updated equivalencies in the same format as before.',
            },
        ])
        return finalOutput['message']['content']
    except httpx.RequestError as e:
        print(f"An HTTP request error occurred: {e}")
    except httpx.HTTPStatusError as e:
        print(f"An HTTP status error occurred: {e.response.status_code} - {e.response.text}")
    except Exception as e:
        print(f"An unexpected error occurred: {e}")

    return None

def save_updated_mapping(final_output):
    # Parse the final output to create a new conversion sheet DataFrame
    updated_mappings = [line.split(":") for line in final_output.strip().split("\n")]
    updated_conversion_sheet = pd.DataFrame(updated_mappings, columns=['Cloud', 'Legacy'])
    updated_conversion_sheet['Cloud'] = updated_conversion_sheet['Cloud'].str.strip()
    updated_conversion_sheet['Legacy'] = updated_conversion_sheet['Legacy'].str.strip()

    # Save to a new Excel file
    updated_conversion_sheet_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/Mapping File/Updated_Vendor.xlsx'
    updated_conversion_sheet.to_excel(updated_conversion_sheet_path, index=False)

    return updated_conversion_sheet, updated_conversion_sheet_path

def transfer_data(updated_conversion_sheet, legacy_sheet_path, cloud_sheet_path, output_path):
    legacy_sheet = pd.read_excel(legacy_sheet_path)
    cloud_sheet = pd.read_excel(cloud_sheet_path)

    conversion_dict = pd.Series(updated_conversion_sheet['Cloud'].values, index=updated_conversion_sheet['Legacy']).to_dict()

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

    cloud_data.to_csv(output_path, index=False)

    print(f"Data transfer complete. Check '{output_path}'")
    print("Transfer Log:")
    print(transfer_log)
    print("Cloud Data Preview:")
    print(cloud_data_preview)

def test_ollama():
    conversion_string, conversion_sheet, legacy_sheet_path, cloud_sheet_path = load_and_prepare_sheets()
    
    cloud_list = conversion_sheet['Cloud'].tolist()
    legacy_list = conversion_sheet['Legacy'].tolist()
    
    final_output = call_ollama_api(conversion_string, cloud_list, legacy_list)
    
    if final_output:
        updated_conversion_sheet, updated_conversion_sheet_path = save_updated_mapping(final_output)
        
        # Use the same paths for legacy and cloud sheets
        output_path = '/Users/krishnasuresh/Desktop/VDart/West Orange - Krishna/updated_po_cloud_sheet.csv'
        
        transfer_data(updated_conversion_sheet, legacy_sheet_path, cloud_sheet_path, output_path)

# Run the test function
test_ollama()
