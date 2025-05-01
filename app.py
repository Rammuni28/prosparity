# # import streamlit as st
# # import os
# # import csv
# # from link_extractor import download_equifax_jsons
# # from data_extractor import generate_excel_from_json

# # # Configuration
# # JSON_DIR = "equifax_jsons"
# # LINKS_TEMP = "temp_links.csv"
# # META_TEMP = "temp_meta.csv"

# # # Ensure JSON directory exists
# # os.makedirs(JSON_DIR, exist_ok=True)

# # # Streamlit setup
# # st.set_page_config(page_title="Bureau Report Processor", layout="wide")
# # st.title("📥 Bureau Report Processor")

# # # Utility to count links in CSV
# # def count_links(csv_path):
# #     with open(csv_path, newline='') as f:
# #         return sum(1 for _ in csv.reader(f))

# # # --- Step 1: Upload and Download JSONs ---
# # links_csv = st.file_uploader("📄 Upload Links CSV (JSON URLs)", type="csv", key="links")
# # if links_csv:
# #     if st.button("📥 Download JSONs", key="download_btn"):
# #         # Save uploaded CSV
# #         with open(LINKS_TEMP, "wb") as f:
# #             f.write(links_csv.getbuffer())
# #         total = count_links(LINKS_TEMP)

# #         # Download with spinner (no hook support)
# #         with st.spinner(f"Downloading {total} JSON files..."):
# #             failed = download_equifax_jsons(LINKS_TEMP)
# #             st.session_state["failed"] = failed or []
# #             st.session_state["download_complete"] = True

# # # Show download results
# # if "download_complete" in st.session_state:
# #     if st.session_state.failed:
# #         st.warning(f"Download finished with {len(st.session_state.failed)} failures.")
# #         st.text_area("Failed URLs/files:", "\n".join(st.session_state.failed), height=150)
# #     else:
# #         st.success("All JSONs downloaded successfully!")

# # # --- Step 2: Process JSONs & Generate Excel ---
# # meta_csv = st.file_uploader("📑 Upload Metadata CSV", type="csv", key="meta")
# # if meta_csv:
# #     if not st.session_state.get("download_complete", False):
# #         st.info("Please download JSONs first.")
# #         st.button("Generate Excel", disabled=True)
# #     else:
# #         if st.button("📊 Generate & Download Excel", key="gen_btn"):
# #             with open(META_TEMP, "wb") as f:
# #                 f.write(meta_csv.getbuffer())

# #             with st.spinner("Generating Excel from JSONs and metadata..."):
# #                 try:
# #                     excel_data = generate_excel_from_json(JSON_DIR, META_TEMP)
# #                 except Exception as e:
# #                     st.error(f"Error during Excel generation: {e}")
# #                     st.stop()

# #             st.success("Excel ready for download!")
# #             st.download_button(
# #                 "📥 Download Excel",
# #                 data=excel_data,
# #                 file_name="bureau_report_output.xlsx",
# #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# #             )
# # import streamlit as st
# # import os
# # import pandas as pd
# # from link_extractor import download_equifax_jsons
# # from data_extractor import generate_excel_from_json

# # # Configuration
# # JSON_DIR = "fax_bureau_reports_streamlit"
# # LINKS_TEMP = "temp_links.csv"
# # META_TEMP = "temp_meta.csv"

# # # Ensure JSON directory exists and clear previous failures
# # os.makedirs(JSON_DIR, exist_ok=True)
# # fail_log = os.path.join(JSON_DIR, "failed_downloads.txt")
# # if os.path.isfile(fail_log):
# #     os.remove(fail_log)

# # # Streamlit setup
# # st.set_page_config(page_title="Bureau Report Processor", layout="wide")
# # st.title("📥 Bureau Report Processor")

# # # --- Step 1: Upload Links CSV and Download JSONs ---
# # links_csv = st.file_uploader(
# #     "📄 Upload Links CSV (columns: 'Bureau Type','Bureau json link','Application ID','Type')",
# #     type="csv", key="links"
# # )
# # if links_csv and st.button("📥 Download JSONs", key="download_btn"):
# #     with open(LINKS_TEMP, "wb") as f:
# #         f.write(links_csv.getbuffer())

# #     # Call link_extractor utility
# #     failed = download_equifax_jsons(LINKS_TEMP, save_folder=JSON_DIR)
# #     st.success("Download complete.")

# #     # Show failures if any
# #     if failed:
# #         st.warning(f"{len(failed)} downloads failed.")
# #         st.text_area("Failed files", "\n".join(failed), height=150)
# #     else:
# #         st.success("All JSONs downloaded successfully!")

# #     st.session_state["download_complete"] = True

# # # --- Step 2: Upload Metadata CSV and Generate Excel ---
# # meta_csv = st.file_uploader("📑 Upload Metadata CSV", type="csv", key="meta")
# # if meta_csv:
# #     if not st.session_state.get("download_complete", False):
# #         st.info("Please download JSONs first.")
# #         st.button("Generate Excel", disabled=True)
# #     else:
# #         if st.button("📊 Generate & Download Excel", key="gen_btn"):
# #             with open(META_TEMP, "wb") as f:
# #                 f.write(meta_csv.getbuffer())

# #             with st.spinner("Generating Excel from JSONs and metadata..."):
# #                 try:
# #                     excel_data = generate_excel_from_json(JSON_DIR, META_TEMP)
# #                 except Exception as e:
# #                     st.error(f"Error during Excel generation: {e}")
# #                     st.stop()

# #             st.success("Excel ready for download!")
# #             st.download_button(
# #                 "📥 Download Excel",
# #                 data=excel_data,
# #                 file_name="bureau_report_output.xlsx",
# #                 mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
# #             )
########
# import streamlit as st
# import os
# import pandas as pd
# import io
# from link_extractor import download_equifax_jsons
# from data_extractor import generate_excel_from_json

# # Configuration
# JSON_DIR = "fax_bureau_reports_streamlit"
# LINKS_TEMP = "temp_links.csv"
# META_TEMP = "temp_meta.csv"
# FAIL_LOG = os.path.join(JSON_DIR, "failed_downloads.txt")

# # Ensure directories and clear old logs
# os.makedirs(JSON_DIR, exist_ok=True)
# if os.path.isfile(FAIL_LOG):
#     os.remove(FAIL_LOG)

# # Streamlit setup
# st.set_page_config(page_title="Bureau Report Processor", layout="wide")
# st.title("📥 Bureau Report Processor")

# # --- Step 1: Download JSONs ---
# links_csv = st.file_uploader("📄 Upload Links CSV (Equifax JSON URLs)", type="csv", key="links")
# if links_csv and st.button("📥 Download JSONs"):
#     with open(LINKS_TEMP, "wb") as f:
#         f.write(links_csv.getbuffer())
#     # Download JSONs, failures logged to FAILED_LOG internally
#     failed = download_equifax_jsons(LINKS_TEMP, JSON_DIR)
#     # Ensure failed is a list, not None
#     if failed is None:
#         failed = []
#     # Or, more succinctly: failed = download_equifax_jsons(LINKS_TEMP, JSON_DIR) or []
#     total = sum(1 for _ in pd.read_csv(LINKS_TEMP).iterrows())
#     succeeded = total - len(failed)
#     # Display counts once
#     st.info(f"✅ Downloaded {succeeded} of {total} JSON files.")
#     if failed:
#         st.warning(f"⚠️ {len(failed)} downloads failed.")
#         st.text_area("Failed JSONs:", "\n".join(failed), height=150)
#     else:
#         st.success("🎉 All JSONs downloaded successfully!")
#     st.session_state["json_done"] = True

# # --- Step 2: Generate Excel ---
# meta_csv = st.file_uploader("📑 Upload Metadata CSV", type="csv", key="meta")
# if meta_csv:
#     if not st.session_state.get("json_done", False):
#         st.info("Please download JSONs first.")
#         st.button("📊 Generate Excel", disabled=True)
#     elif st.button("📊 Generate Excel"):
#         with open(META_TEMP, "wb") as f:
#             f.write(meta_csv.getbuffer())
#         with st.spinner("Generating Excel from JSONs and metadata..."):
#             try:
#                 result = generate_excel_from_json(JSON_DIR, META_TEMP)
#                 # Handle BytesIO
#                 if isinstance(result, io.BytesIO):
#                     excel_bytes = result.getvalue()
#                 # Raw bytes
#                 elif isinstance(result, (bytes, bytearray)):
#                     excel_bytes = result
#                 # Filepath
#                 elif isinstance(result, str) and os.path.isfile(result):
#                     with open(result, "rb") as ef:
#                         excel_bytes = ef.read()
#                     st.info(f"Excel saved at: {result}")
#                 # DataFrame: convert
#                 elif isinstance(result, pd.DataFrame):
#                     buf = io.BytesIO()
#                     with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
#                         result.to_excel(writer, index=False)
#                     excel_bytes = buf.getvalue()
#                     st.info("Converted DataFrame to Excel.")
#                 else:
#                     raise TypeError(f"Unexpected return type: {type(result)}")
#             except Exception as e:
#                 st.error(f"Error during Excel generation: {e}")
#                 st.stop()
#         st.success("Excel generated successfully!")
#         st.download_button(
#             "📥 Download Excel",
#             data=excel_bytes,
#             file_name="bureau_report_output.xlsx",
#             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
#         )
####

import streamlit as st
import os
import pandas as pd
import io
from link_extractor import download_equifax_jsons
from data_extractor import generate_excel_from_json

# Configuration
JSON_DIR = "fax_bureau_reports_streamlit"
LINKS_TEMP = "temp_links.csv"
META_TEMP = "temp_meta.csv"
FAIL_LOG = os.path.join(JSON_DIR, "failed_downloads.txt")

# Ensure directories and clear old logs
os.makedirs(JSON_DIR, exist_ok=True)
if os.path.isfile(FAIL_LOG):
    os.remove(FAIL_LOG)

# Streamlit setup
st.set_page_config(page_title="Bureau Report Processor", layout="wide")
st.title("📥 Bureau Report Processor")

# --- Step 1: Download JSONs ---
links_csv = st.file_uploader("📄 Upload Links CSV (Equifax JSON URLs)", type="csv", key="links")
if links_csv and st.button("📥 Download JSONs"):
    with open(LINKS_TEMP, "wb") as f:
        f.write(links_csv.getbuffer())
    with st.spinner("Downloading JSON files..."):
        failed = download_equifax_jsons(LINKS_TEMP, JSON_DIR)
        if failed is None:
            failed = []
        total = sum(1 for _ in pd.read_csv(LINKS_TEMP).iterrows())
        succeeded = total - len(failed)
    st.info(f"✅ Downloaded {succeeded} of {total} JSON files.")
    if failed:
        st.warning(f"⚠️ {len(failed)} downloads failed.")
        st.text_area("Failed JSONs:", "\n".join(failed), height=150)
    else:
        st.success("🎉 All JSONs downloaded successfully!")
    st.session_state["json_done"] = True

# --- Step 2: Generate Excel ---
meta_csv = st.file_uploader("📑 Upload Metadata CSV", type="csv", key="meta")
if meta_csv:
    if not st.session_state.get("json_done", False):
        st.info("Please download JSONs first.")
        st.button("📊 Generate Excel", disabled=True)
    elif st.button("📊 Generate Excel"):
        with open(META_TEMP, "wb") as f:
            f.write(meta_csv.getbuffer())
        with st.spinner("Generating Excel from JSONs and metadata..."):
            try:
                result = generate_excel_from_json(JSON_DIR, META_TEMP)
                # Handle BytesIO
                if isinstance(result, io.BytesIO):
                    excel_bytes = result.getvalue()
                # Raw bytes
                elif isinstance(result, (bytes, bytearray)):
                    excel_bytes = result
                # Filepath
                elif isinstance(result, str) and os.path.isfile(result):
                    with open(result, "rb") as ef:
                        excel_bytes = ef.read()
                    st.info(f"Excel saved at: {result}")
                # DataFrame: convert
                elif isinstance(result, pd.DataFrame):
                    buf = io.BytesIO()
                    with pd.ExcelWriter(buf, engine="xlsxwriter") as writer:
                        result.to_excel(writer, index=False)
                    excel_bytes = buf.getvalue()
                    st.info("Converted DataFrame to Excel.")
                else:
                    raise TypeError(f"Unexpected return type: {type(result)}")
            except Exception as e:
                st.error(f"Error during Excel generation: {e}")
                st.stop()
        st.success("Excel generated successfully!")
        st.download_button(
            "📥 Download Excel",
            data=excel_bytes,
            file_name="bureau_report_output.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
