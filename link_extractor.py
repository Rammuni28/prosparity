# import os
# import pandas as pd
# import requests

# def download_equifax_jsons(csv_file_path, save_folder="fax_bureau_reports_streamlit"):
#     df = pd.read_csv(csv_file_path)
#     os.makedirs(save_folder, exist_ok=True)
#     failed_files = []

#     def download_json(url, filename):
#         try:
#             response = requests.get(url, timeout=15)
#             if response.status_code == 200:
#                 with open(os.path.join(save_folder, filename), "w", encoding="utf-8") as file:
#                     file.write(response.text)
#                 print(f"Downloaded: {filename}")
#             else:
#                 print(f"Failed to download {filename}, Status Code: {response.status_code}")
#                 failed_files.append(filename)
#         except Exception as e:
#             print(f"Error downloading {filename}: {e}")
#             failed_files.append(filename)

#     for index, row in df[df["Bureau Type"] == "Equifax"].iterrows():
#         json_url = row["Bureau json link"]
#         app_id = row["Application ID"]
#         file_type = row["Type"].replace(" ", "_")
#         filename = f"{app_id}_{file_type}.json"
#         download_json(json_url, filename)

#     return failed_files

##############
import os
import pandas as pd
import requests

def download_equifax_jsons(csv_file_path, save_folder="fax_bureau_reports_streamlit"):
    df = pd.read_csv(csv_file_path)
    os.makedirs(save_folder, exist_ok=True)

    def download_json(url, filename):
        try:
            response = requests.get(url, timeout=15)
            if response.status_code == 200:
                with open(os.path.join(save_folder, filename), "w", encoding="utf-8") as file:
                    file.write(response.text)
                print(f"Downloaded: {filename}")
            else:
                print(f"Failed to download {filename}, Status Code: {response.status_code}")
                with open("failed_downloads.txt", "a") as fail_log:
                    fail_log.write(filename + "\n")
        except Exception as e:
            print(f"Error downloading {filename}: {e}")
            with open("failed_downloads.txt", "a") as fail_log:
                fail_log.write(filename + "\n")

    for index, row in df[df["Bureau Type"] == "Equifax"].iterrows():
        json_url = row["Bureau json link"]
        app_id = row["Application ID"]
        file_type = row["Type"].replace(" ", "_")
        filename = f"{app_id}_{file_type}.json"
        download_json(json_url, filename)
######
