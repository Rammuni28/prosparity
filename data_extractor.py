import os
import glob
import json
import pandas as pd
from datetime import datetime
from dateutil.relativedelta import relativedelta
from io import BytesIO

def generate_excel_from_json(json_folder_path, metadata_csv_path):
    application_dates = read_application_date_csv(metadata_csv_path)
    records = []

    for file in os.listdir(json_folder_path):
        if file.endswith(".json"):
            path = os.path.join(json_folder_path, file)
            record = parse_json_file(path, override_date=None)
            if record:
                records.append(record)

    final_df = pd.DataFrame(records)

    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        final_df.to_excel(writer, index=False)
    output.seek(0)

    return output


def safe_float(value):
    """Convert a value to float safely; if conversion fails, return 0."""
    try:
        return float(value)
    except (ValueError, TypeError):
        return 0

def read_application_date_csv(csv_path):
    """
    Reads the CSV file and builds a dictionary mapping Application ID to Date Of Application.
    Assumes the CSV has columns: "Application ID" and "Date Of Application".
    """
    try:
        df = pd.read_csv(csv_path)
        mapping = dict(zip(df["Applicant ID"], df["Date Of Application"]))
        return mapping
    except Exception as e:
        print(f"Error reading CSV file {csv_path}: {str(e)}")
        return {}

def parse_json_file(file_path, override_date=None):
    """
    Parse JSON file and extract required details.
    override_date: (optional) if provided, used instead of computed bureau pull date
                   for all time-difference calculations.
    """
    # Extract Application ID and Type from filename (e.g., "PROSAPP240917000004_Coborrower.json")
    filename = os.path.basename(file_path)
    parts = filename.replace(".json", "").split("_")
    application_id = parts[0]
    application_type = parts[1] if len(parts) > 1 else ""

    try:
        with open(file_path, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except (json.JSONDecodeError, IOError) as e:
        print(f"Error reading {file_path}: {str(e)}")
        return None
    
    report = data.get("equifaxReport", {})
    id_info = report.get("IDAndContactInfo", {})
    personal = id_info.get("PersonalInfo", {})
    # identity = id_info.get("IDAndContactInfo", {}).get("IdentityInfo", {})
    identity = report.get("IDAndContactInfo", {}) \
                 .get("IdentityInfo", {})


    # Personal Information
    name_info = personal.get("Name", {})
    consumer_name = name_info.get("FullName", "").strip()

    # Date handling with validation for DOB
    dob = personal.get("DateOfBirth", "")
    try:
        dob_datetime = pd.to_datetime(dob, errors='coerce')
        formatted_dob = dob_datetime.strftime("%Y-%m-%d %H:%M:%S") if not pd.isnull(dob_datetime) else ""
    except Exception:
        formatted_dob = ""

    # Account Details Extraction
    accounts = report.get("RetailAccountDetails", [])
    summary = report.get("RetailAccountsSummary", {})
    
    summary_total_past_due = safe_float(summary.get("TotalPastDue", 0))
    summary_total_balance = safe_float(summary.get("TotalBalanceAmount", 0))
    
    tradeline_total_past_due = 0
    tradeline_total_balance = 0
    
    institutions, account_types, ownership_types = [], [], []
    
    # Initialize buckets for DPD and other counts
    dpd_01_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    dpd_30_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    dpd_61_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    dpd_91_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    dpd_180_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    dpd_181_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    
    impaired_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    impaired_pwos_codes = ["PWOS", "SPM", "RHM", "SET", "SF", "DBT", "FPD", "WDF", "WOF", "SWDW", "WDWO", "SFWO", "LOSS", "SFR", "SFWD"]
    impaired_pwos_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}
    
    substandard_counts = {key: 0 for key in ["last_1", "last_2", "last_3", "last_6", "last_12", "last_24", "last_36", "ever"]}

    # Compute bureau pull date from accounts (maximum DateReported)
    pull_dates = []
    for acc in accounts:
        dr = acc.get("DateReported", "")
        try:
            pull_dates.append(pd.to_datetime(dr))
        except Exception:
            continue
    computed_date = max(pull_dates) if pull_dates else None

    # Use override_date (CSV "Date Of Application") if provided, else use computed_date
    if override_date:
        bureau_pull_date = pd.to_datetime(override_date, errors='coerce')
    else:
        bureau_pull_date = computed_date

    # Loop through each account to update aggregates and DPD counts using the selected bureau_pull_date
    for acc in accounts:
        institutions.append(acc.get("Institution", ""))
        account_types.append(acc.get("AccountType", ""))
        ownership_types.append(acc.get("OwnershipType", ""))
        
        past_due_amount = safe_float(acc.get("PastDueAmount", 0))
        tradeline_total_past_due += past_due_amount
        
        balance_amount = safe_float(acc.get("Balance", 0))
        tradeline_total_balance += balance_amount
        
        histories = acc.get("History48Months", [])
        for rec in histories:
            status = rec.get("PaymentStatus", "").strip()
            try:
                rec_date = pd.to_datetime(rec.get("key", ""), format="%m-%y")
            except Exception:
                continue
            if bureau_pull_date and rec_date <= bureau_pull_date:
                diff = relativedelta(bureau_pull_date, rec_date)
                months_diff = diff.years * 12 + diff.months
                # DPD 01+ counts
                if status == "01+":
                    if months_diff < 1:
                        dpd_01_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_01_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_01_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_01_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_01_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_01_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_01_counts["last_36"] += 1
                    dpd_01_counts["ever"] += 1
                # DPD 30+ counts
                if status == "30+":
                    if months_diff < 1:
                        dpd_30_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_30_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_30_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_30_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_30_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_30_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_30_counts["last_36"] += 1
                    dpd_30_counts["ever"] += 1
                # DPD 61+ counts
                if status == "61+":
                    if months_diff < 1:
                        dpd_61_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_61_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_61_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_61_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_61_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_61_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_61_counts["last_36"] += 1
                    dpd_61_counts["ever"] += 1
                # DPD 91+ counts
                if status in ("91+", "121+"):
                    if months_diff < 1:
                        dpd_91_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_91_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_91_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_91_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_91_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_91_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_91_counts["last_36"] += 1
                    dpd_91_counts["ever"] += 1
                # DPD 180+ counts
                if status == "180DPD":
                    if months_diff < 1:
                        dpd_180_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_180_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_180_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_180_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_180_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_180_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_180_counts["last_36"] += 1
                    dpd_180_counts["ever"] += 1
                # DPD 181+ counts
                if status == "181+":
                    if months_diff < 1:
                        dpd_181_counts["last_1"] += 1
                    if months_diff < 2:
                        dpd_181_counts["last_2"] += 1
                    if months_diff < 3:
                        dpd_181_counts["last_3"] += 1
                    if months_diff < 6:
                        dpd_181_counts["last_6"] += 1
                    if months_diff < 12:
                        dpd_181_counts["last_12"] += 1
                    if months_diff < 24:
                        dpd_181_counts["last_24"] += 1
                    if months_diff < 36:
                        dpd_181_counts["last_36"] += 1
                    dpd_181_counts["ever"] += 1
                # Impaired Entry counts
                if status == "Impaired":
                    if months_diff < 1:
                        impaired_counts["last_1"] += 1
                    if months_diff < 2:
                        impaired_counts["last_2"] += 1
                    if months_diff < 3:
                        impaired_counts["last_3"] += 1
                    if months_diff < 6:
                        impaired_counts["last_6"] += 1
                    if months_diff < 12:
                        impaired_counts["last_12"] += 1
                    if months_diff < 24:
                        impaired_counts["last_24"] += 1
                    if months_diff < 36:
                        impaired_counts["last_36"] += 1
                    impaired_counts["ever"] += 1
                # Impaired PWOS-type counts
                if status in impaired_pwos_codes:
                    if months_diff < 1:
                        impaired_pwos_counts["last_1"] += 1
                    if months_diff < 2:
                        impaired_pwos_counts["last_2"] += 1
                    if months_diff < 3:
                        impaired_pwos_counts["last_3"] += 1
                    if months_diff < 6:
                        impaired_pwos_counts["last_6"] += 1
                    if months_diff < 12:
                        impaired_pwos_counts["last_12"] += 1
                    if months_diff < 24:
                        impaired_pwos_counts["last_24"] += 1
                    if months_diff < 36:
                        impaired_pwos_counts["last_36"] += 1
                    impaired_pwos_counts["ever"] += 1
                # Sub Standard counts
                if status == "SUB":
                    if months_diff < 1:
                        substandard_counts["last_1"] += 1
                    if months_diff < 2:
                        substandard_counts["last_2"] += 1
                    if months_diff < 3:
                        substandard_counts["last_3"] += 1
                    if months_diff < 6:
                        substandard_counts["last_6"] += 1
                    if months_diff < 12:
                        substandard_counts["last_12"] += 1
                    if months_diff < 24:
                        substandard_counts["last_24"] += 1
                    if months_diff < 36:
                        substandard_counts["last_36"] += 1
                    substandard_counts["ever"] += 1

    # Enquiry Details extraction
    enquiry = report.get("EnquirySummary", {})
    total_enquiry = enquiry.get("Total", "")
    enquiry_past30 = enquiry.get("Past30Days", "")
    enquiry_past12 = enquiry.get("Past12Months", "")
    enquiry_past24 = enquiry.get("Past24Months", "")

    def create_indexed_dict(items):
        return {idx+1: item for idx, item in enumerate(sorted(set(filter(None, items))))}

    return {
        "Application ID": application_id,
        "Type": application_type,
        "consumer_name": consumer_name,
        "gender": personal.get("Gender", "").strip(),
        "dob": formatted_dob,
        "age": personal.get("Age", {}).get("Age", "").strip(),
        "total_income": personal.get("TotalIncome", "").strip(),
        "pan": (identity.get("PANId", [{}])[0]
               .get("IdNumber", "")
               .strip()),
        # "pan": identity.get("PANId", [{}])[-1].get("IdNumber", "").strip(),
        "address": id_info.get("AddressInfo", [{}])[-1].get("Address", "").strip(),
        "state": id_info.get("AddressInfo", [{}])[-1].get("State", "").strip(),
        "mobile": id_info.get("PhoneInfo", [{}])[-1].get("Number", "").strip(),
        "bureau_score": str(report.get("ScoreDetails", [{}])[0].get("Value", "")),
        "institutions": create_indexed_dict(institutions),
        "account_types": create_indexed_dict(account_types),
        "ownership_types": create_indexed_dict(ownership_types),
        "summary_total_past_due": summary_total_past_due,
        "tradeline_total_past_due": tradeline_total_past_due,
        "summary_total_balance": summary_total_balance,
        "tradeline_total_balance": tradeline_total_balance,
        "dpd_01_counts": dpd_01_counts,
        "dpd_30_counts": dpd_30_counts,
        "dpd_61_counts": dpd_61_counts,
        "dpd_91_counts": dpd_91_counts,
        "dpd_180_counts": dpd_180_counts,
        "dpd_181_counts": dpd_181_counts,
        "impaired_counts": impaired_counts,
        "impaired_pwos_counts": impaired_pwos_counts,
        "substandard_counts": substandard_counts,
        "total_enquiry": total_enquiry,
        "enquiry_past30": enquiry_past30,
        "enquiry_past12": enquiry_past12,
        "enquiry_past24": enquiry_past24,
        "bureau_pull_date": bureau_pull_date.strftime("%Y-%m-%d") if bureau_pull_date else ""
    }

def generate_excel_output(json_folder, output_file, csv_app_file):
    """Process JSON files, override bureau pull date with CSV Date Of Application, and generate Excel report."""
    results = []
    # Read CSV mapping Application ID to Date Of Application
    app_date_mapping = read_application_date_csv(csv_app_file)

    for json_file in glob.glob(os.path.join(json_folder, "*.json")):
        # Extract application id from filename (already done in parse_json_file)
        # Get CSV override date using application id
        base_name = os.path.basename(json_file)
        application_id = base_name.split('_')[0]
        override_date = app_date_mapping.get(application_id, None)
        
        file_data = parse_json_file(json_file, override_date=override_date)
        if not file_data:
            continue

        dpd01 = file_data["dpd_01_counts"]
        dpd30 = file_data["dpd_30_counts"]
        dpd61 = file_data["dpd_61_counts"]
        dpd91 = file_data["dpd_91_counts"]
        dpd180 = file_data["dpd_180_counts"]
        dpd181 = file_data["dpd_181_counts"]
        impaired = file_data["impaired_counts"]
        impaired_pwos = file_data["impaired_pwos_counts"]
        substandard = file_data["substandard_counts"]

        row = {
            "Application ID": file_data["Application ID"],
            "Type": file_data["Type"],
            "Application Current Status": "From Processed Application Report",
            "Consumer Name": file_data["consumer_name"],
            "Gender": file_data["gender"],
            "PAN": file_data["pan"],
            "Address": file_data["address"],
            "DOB": file_data["dob"],
            "State": file_data["state"],
            "Mobile": file_data["mobile"],
            "Age": file_data["age"],
            "Total Income": file_data["total_income"],
            "Bureau Score": file_data["bureau_score"],
            "Account Institutions": str(file_data["institutions"]),
            "Account AccountTypes": str(file_data["account_types"]),
            "Account OwnershipTypes": str(file_data["ownership_types"]),
            "Summary Total Past Due": file_data["summary_total_past_due"],
            "Tradeline Total Past Due": file_data["tradeline_total_past_due"],
            "Summary Total Balance": file_data["summary_total_balance"],
            "Tradeline Total Balance": file_data["tradeline_total_balance"],
            # (Yahan pe aap apne DPD, Impaired, Substandard counts ke columns add kar sakte hain;
            # maine niche se existing logic ko as-is rehne diya hai)
            "Count of 1-30DPD in last 1 month": dpd01["last_1"],
            "Count of 1-30DPD in last 2 months": dpd01["last_2"],
            "Count of 1-30DPD in last 3 months": dpd01["last_3"],
            "Count of 1-30DPD in last 6 months": dpd01["last_6"],
            "Count of 1-30DPD in last 12 months": dpd01["last_12"],
            "Count of 1-30DPD in last 24 months": dpd01["last_24"],
            "Count of 1-30DPD in last 36 months": dpd01["last_36"],
            "Count of 1-30DPD ever": dpd01["ever"],
            "Number of 01+ in the last 1 month from the bureau pull date": dpd01["last_1"],
            "Number of 01+ in the last 2 months from the bureau pull date": dpd01["last_2"],
            "Number of 01+ in the last 3 months from the bureau pull date": dpd01["last_3"],
            "Number of 01+ in the last 6 months from the bureau pull date": dpd01["last_6"],
            "Number of 01+ in the last 12 months from the bureau pull date": dpd01["last_12"],
            "Number of 01+ in the last 24 months from the bureau pull date": dpd01["last_24"],
            "Number of 01+ in the last 36 months from the bureau pull date": dpd01["last_36"],
            "Number of 01+ ever": dpd01["ever"],
            "Count of 31-60DPD in last 1 month": dpd30["last_1"],
            "Count of 31-60DPD in last 2 months": dpd30["last_2"],
            "Count of 31-60DPD in last 3 months": dpd30["last_3"],
            "Count of 31-60DPD in last 6 months": dpd30["last_6"],
            "Count of 31-60DPD in last 12 months": dpd30["last_12"],
            "Count of 31-60DPD in last 24 months": dpd30["last_24"],
            "Count of 31-60DPD in last 36 months": dpd30["last_36"],
            "Count of 31-60DPD ever": dpd30["ever"],
            "Number of 30+ in the last 1 month from the bureau pull date": dpd30["last_1"],
            "Number of 30+ in the last 2 months from the bureau pull date": dpd30["last_2"],
            "Number of 30+ in the last 3 months from the bureau pull date": dpd30["last_3"],
            "Number of 30+ in the last 6 months from the bureau pull date": dpd30["last_6"],
            "Number of 30+ in the last 12 months from the bureau pull date": dpd30["last_12"],
            "Number of 30+ in the last 24 months from the bureau pull date": dpd30["last_24"],
            "Number of 30+ in the last 36 months from the bureau pull date": dpd30["last_36"],
            "Number of 30+ ever": dpd30["ever"],
            "Count of 61-90DPD in last 1 month": dpd61["last_1"],
            "Count of 61-90DPD in last 2 months": dpd61["last_2"],
            "Count of 61-90DPD in last 3 months": dpd61["last_3"],
            "Count of 61-90DPD in last 6 months": dpd61["last_6"],
            "Count of 61-90DPD in last 12 months": dpd61["last_12"],
            "Count of 61-90DPD in last 24 months": dpd61["last_24"],
            "Count of 61-90DPD in last 36 months": dpd61["last_36"],
            "Count of 61-90DPD ever": dpd61["ever"],
            "Number of 61+ in the last 1 month from the bureau pull date": dpd61["last_1"],
            "Number of 61+ in the last 2 months from the bureau pull date": dpd61["last_2"],
            "Number of 61+ in the last 3 months from the bureau pull date": dpd61["last_3"],
            "Number of 61+ in the last 6 months from the bureau pull date": dpd61["last_6"],
            "Number of 61+ in the last 12 months from the bureau pull date": dpd61["last_12"],
            "Number of 61+ in the last 24 months from the bureau pull date": dpd61["last_24"],
            "Number of 61+ in the last 36 months from the bureau pull date": dpd61["last_36"],
            "Number of 61+ ever": dpd61["ever"],
            "Count of 91-179DPD in last 1 month": dpd91["last_1"],
            "Count of 91-179DPD in last 2 months": dpd91["last_2"],
            "Count of 91-179DPD in last 3 months": dpd91["last_3"],
            "Count of 91-179DPD in last 6 months": dpd91["last_6"],
            "Count of 91-179DPD in last 12 months": dpd91["last_12"],
            "Count of 91-179DPD in last 24 months": dpd91["last_24"],
            "Count of 91-179DPD in last 36 months": dpd91["last_36"],
            "Count of 91-179DPD ever": dpd91["ever"],
            "Number of 91+ and 121+ in the last 1 month from the bureau pull date": dpd91["last_1"],
            "Number of 91+ and 121+ in the last 2 months from the bureau pull date": dpd91["last_2"],
            "Number of 91+ and 121+ in the last 3 months from the bureau pull date": dpd91["last_3"],
            "Number of 91+ and 121+ in the last 6 months from the bureau pull date": dpd91["last_6"],
            "Number of 91+ and 121+ in the last 12 months from the bureau pull date": dpd91["last_12"],
            "Number of 91+ and 121+ in the last 24 months from the bureau pull date": dpd91["last_24"],
            "Number of 91+ and 121+ in the last 36 months from the bureau pull date": dpd91["last_36"],
            "Number of 91+ and 121+ ever": dpd91["ever"],
            "Count of 180DPD in last 1 month": dpd180["last_1"],
            "Count of 180DPD in last 2 months": dpd180["last_2"],
            "Count of 180DPD in last 3 months": dpd180["last_3"],
            "Count of 180DPD in last 6 months": dpd180["last_6"],
            "Count of 180DPD in last 12 months": dpd180["last_12"],
            "Count of 180DPD in last 24 months": dpd180["last_24"],
            "Count of 180DPD in last 36 months": dpd180["last_36"],
            "Count of 180DPD ever": dpd180["ever"],
            "Number of 181+ in the last 1 month from the bureau pull date": dpd181["last_1"],
            "Number of 181+ in the last 2 months from the bureau pull date": dpd181["last_2"],
            "Number of 181+ in the last 3 months from the bureau pull date": dpd181["last_3"],
            "Number of 181+ in the last 6 months from the bureau pull date": dpd181["last_6"],
            "Number of 181+ in the last 12 months from the bureau pull date": dpd181["last_12"],
            "Number of 181+ in the last 24 months from the bureau pull date": dpd181["last_24"],
            "Number of 181+ in the last 36 months from the bureau pull date": dpd181["last_36"],
            "Number of 181+ ever": dpd181["ever"],
            "Count of Impaired Entry in last 1 month": impaired["last_1"],
            "Count of Impaired Entry in last 2 months": impaired["last_2"],
            "Count of Impaired Entry in last 3 months": impaired["last_3"],
            "Count of Impaired Entry in last 6 months": impaired["last_6"],
            "Count of Impaired Entry in last 12 months": impaired["last_12"],
            "Count of Impaired Entry in last 24 months": impaired["last_24"],
            "Count of Impaired Entry in last 36 months": impaired["last_36"],
            "Count of Impaired Entry ever": impaired["ever"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 1 month from the bureau pull date": impaired_pwos["last_1"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 2 months from the bureau pull date": impaired_pwos["last_2"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 3 months from the bureau pull date": impaired_pwos["last_3"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 6 months from the bureau pull date": impaired_pwos["last_6"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 12 months from the bureau pull date": impaired_pwos["last_12"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 24 months from the bureau pull date": impaired_pwos["last_24"],
            "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 36 months from the bureau pull date": impaired_pwos["last_36"],
            "Number of PWOS SPM RHM SET SF DBT WDF WOF SWDW WDWO SFWO LOSS SFR SFWD ever": impaired_pwos["ever"],
            "Count of Sub standard in last 1 month": substandard["last_1"],
            "Count of Sub standard in last 2 months": substandard["last_2"],
            "Count of Sub standard in last 3 months": substandard["last_3"],
            "Count of Sub standard in last 6 months": substandard["last_6"],
            "Count of Sub standard in last 12 months": substandard["last_12"],
            "Count of Sub standard in last 24 months": substandard["last_24"],
            "Count of Sub standard in last 36 months": substandard["last_36"],
            "Count of Sub standard ever": substandard["ever"],
            "Number of SUB in the last 1 month from the bureau pull date": substandard["last_1"],
            "Number of SUB in the last 2 months from the bureau pull date": substandard["last_2"],
            "Number of SUB in the last 3 months from the bureau pull date": substandard["last_3"],
            "Number of SUB in the last 6 months from the bureau pull date": substandard["last_6"],
            "Number of SUB in the last 12 months from the bureau pull date": substandard["last_12"],
            "Number of SUB in the last 24 months from the bureau pull date": substandard["last_24"],
            "Number of SUB in the last 36 months from the bureau pull date": substandard["last_36"],
            "Number of SUB ever": substandard["ever"],
            "Total Enquiry": file_data["total_enquiry"],
            "Enquiry Past30Days": file_data["enquiry_past30"],
            "Enquiry Past12Months": file_data["enquiry_past12"],
            "Enquiry Past24Months": file_data["enquiry_past24"],
            # "Bureau Pull Date" will now show the CSV "Date Of Application"
            "Bureau Pull Date": file_data["bureau_pull_date"]
        }
        results.append(row)

    columns = [
        "Application ID",
        "Type",
        "Application Current Status", "Consumer Name", "Gender",
        "PAN", "Address", "DOB", "State", "Mobile", "Age", "Total Income",
        "Bureau Score", "Account Institutions", "Account AccountTypes",
        "Account OwnershipTypes", "Summary Total Past Due", "Tradeline Total Past Due",
        "Summary Total Balance", "Tradeline Total Balance",
        # (Include DPD counts and other columns in your desired order)
        "Count of 1-30DPD in last 1 month", "Count of 1-30DPD in last 2 months",
        "Count of 1-30DPD in last 3 months", "Count of 1-30DPD in last 6 months",
        "Count of 1-30DPD in last 12 months", "Count of 1-30DPD in last 24 months",
        "Count of 1-30DPD in last 36 months", "Count of 1-30DPD ever",
        "Number of 01+ in the last 1 month from the bureau pull date",
        "Number of 01+ in the last 2 months from the bureau pull date",
        "Number of 01+ in the last 3 months from the bureau pull date",
        "Number of 01+ in the last 6 months from the bureau pull date",
        "Number of 01+ in the last 12 months from the bureau pull date",
        "Number of 01+ in the last 24 months from the bureau pull date",
        "Number of 01+ in the last 36 months from the bureau pull date",
        "Number of 01+ ever",
        "Count of 31-60DPD in last 1 month", "Count of 31-60DPD in last 2 months",
        "Count of 31-60DPD in last 3 months", "Count of 31-60DPD in last 6 months",
        "Count of 31-60DPD in last 12 months", "Count of 31-60DPD in last 24 months",
        "Count of 31-60DPD in last 36 months", "Count of 31-60DPD ever",
        "Number of 30+ in the last 1 month from the bureau pull date",
        "Number of 30+ in the last 2 months from the bureau pull date",
        "Number of 30+ in the last 3 months from the bureau pull date",
        "Number of 30+ in the last 6 months from the bureau pull date",
        "Number of 30+ in the last 12 months from the bureau pull date",
        "Number of 30+ in the last 24 months from the bureau pull date",
        "Number of 30+ in the last 36 months from the bureau pull date",
        "Number of 30+ ever",
        "Count of 61-90DPD in last 1 month", "Count of 61-90DPD in last 2 months",
        "Count of 61-90DPD in last 3 months", "Count of 61-90DPD in last 6 months",
        "Count of 61-90DPD in last 12 months", "Count of 61-90DPD in last 24 months",
        "Count of 61-90DPD in last 36 months", "Count of 61-90DPD ever",
        "Number of 61+ in the last 1 month from the bureau pull date",
        "Number of 61+ in the last 2 months from the bureau pull date",
        "Number of 61+ in the last 3 months from the bureau pull date",
        "Number of 61+ in the last 6 months from the bureau pull date",
        "Number of 61+ in the last 12 months from the bureau pull date",
        "Number of 61+ in the last 24 months from the bureau pull date",
        "Number of 61+ in the last 36 months from the bureau pull date",
        "Number of 61+ ever",
        "Count of 91-179DPD in last 1 month", "Count of 91-179DPD in last 2 months",
        "Count of 91-179DPD in last 3 months", "Count of 91-179DPD in last 6 months",
        "Count of 91-179DPD in last 12 months", "Count of 91-179DPD in last 24 months",
        "Count of 91-179DPD in last 36 months", "Count of 91-179DPD ever",
        "Number of 91+ and 121+ in the last 1 month from the bureau pull date",
        "Number of 91+ and 121+ in the last 2 months from the bureau pull date",
        "Number of 91+ and 121+ in the last 3 months from the bureau pull date",
        "Number of 91+ and 121+ in the last 6 months from the bureau pull date",
        "Number of 91+ and 121+ in the last 12 months from the bureau pull date",
        "Number of 91+ and 121+ in the last 24 months from the bureau pull date",
        "Number of 91+ and 121+ in the last 36 months from the bureau pull date",
        "Number of 91+ and 121+ ever",
        "Count of 180DPD in last 1 month", "Count of 180DPD in last 2 months",
        "Count of 180DPD in last 3 months", "Count of 180DPD in last 6 months",
        "Count of 180DPD in last 12 months", "Count of 180DPD in last 24 months",
        "Count of 180DPD in last 36 months", "Count of 180DPD ever",
        "Number of 181+ in the last 1 month from the bureau pull date",
        "Number of 181+ in the last 2 months from the bureau pull date",
        "Number of 181+ in the last 3 months from the bureau pull date",
        "Number of 181+ in the last 6 months from the bureau pull date",
        "Number of 181+ in the last 12 months from the bureau pull date",
        "Number of 181+ in the last 24 months from the bureau pull date",
        "Number of 181+ in the last 36 months from the bureau pull date",
        "Number of 181+ ever",
        "Count of Impaired Entry in last 1 month", "Count of Impaired Entry in last 2 months",
        "Count of Impaired Entry in last 3 months", "Count of Impaired Entry in last 6 months",
        "Count of Impaired Entry in last 12 months", "Count of Impaired Entry in last 24 months",
        "Count of Impaired Entry in last 36 months", "Count of Impaired Entry ever",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 1 month from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 2 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 3 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 6 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 12 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 24 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD in the last 36 months from the bureau pull date",
        "Number of PWOS SPM RHM SET SF DBT FPD WDF WOF SWDW WDWO SFWO LOSS SFR SFWD ever",
        "Count of Sub standard in last 1 month", "Count of Sub standard in last 2 months",
        "Count of Sub standard in last 3 months", "Count of Sub standard in last 6 months",
        "Count of Sub standard in last 12 months", "Count of Sub standard in last 24 months",
        "Count of Sub standard in last 36 months", "Count of Sub standard ever",
        "Number of SUB in the last 1 month from the bureau pull date",
        "Number of SUB in the last 2 months from the bureau pull date",
        "Number of SUB in the last 3 months from the bureau pull date",
        "Number of SUB in the last 6 months from the bureau pull date",
        "Number of SUB in the last 12 months from the bureau pull date",
        "Number of SUB in the last 24 months from the bureau pull date",
        "Number of SUB in the last 36 months from the bureau pull date",
        "Number of SUB ever",
        "Total Enquiry", "Enquiry Past30Days", "Enquiry Past12Months", "Enquiry Past24Months",
        "Bureau Pull Date"
    ]

    df = pd.DataFrame(results, columns=columns)
    df.to_excel(output_file, index=False)
    print(f"Successfully generated report with {len(results)} rows at {output_file}")

# if __name__ == "__main__":
#     JSON_FOLDER = "equifax_bureau_reports_apr/"
#     OUTPUT_FILE = "humpy_tuntun_bureau_report.xlsx"
#     CSV_APP_FILE = "/Users/rammunidiwan/proparity/processed_app_report_2024-09-01_to_2025-04-30.csv"  # CSV with Application ID and Date Of Application
#     generate_excel_output(JSON_FOLDER, OUTPUT_FILE, CSV_APP_FILE)
