import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import requests
import xml.etree.ElementTree as ET
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from datetime import datetime
from dotenv import load_dotenv
import os
import webbrowser

# Load the API key from .env file
load_dotenv()
api_key = os.getenv('ABN_API_KEY')

def fetch_abn_details(abn, api_key):
    abn_clean = abn.replace(" ", "")  # Remove spaces from ABN
    url = f"https://abr.business.gov.au/abrxmlsearch/ABRXMLSearch.asmx/SearchByABNv201408?includeHistoricalDetails=Y&authenticationGuid={api_key}&searchString={abn_clean}"
    try:
        response = requests.get(url)
        if response.status_code == 200:
            return response.content
    except Exception as e:
        print(f"Error fetching details for ABN {abn}: {e}")
    return None

def parse_abn_xml(xml_content):
    if xml_content is None:
        return None
    
    try:
        root = ET.fromstring(xml_content)
        namespaces = {'abr': 'http://abr.business.gov.au/ABRXMLSearch/'}
        
        business_entity = root.find('.//abr:businessEntity201408', namespaces)
        if business_entity is None:
            return None

        # DGR and Charity Information
        dgr = business_entity.find('.//abr:dgrEndorsement', namespaces)
        charity_types = business_entity.findall('.//abr:charityType', namespaces)
        endorsements = business_entity.findall('.//abr:taxConcessionCharityEndorsement', namespaces)
        acnc_registration = business_entity.find('.//abr:ACNCRegistration', namespaces)
        
        dgr_info = {
            'DGR Endorsement': dgr.findtext('.//abr:entityEndorsement', namespaces=namespaces, default='N/A').strip() if dgr else 'N/A',
            'DGR Start Date': dgr.findtext('.//abr:endorsedFrom', namespaces=namespaces, default='N/A').strip() if dgr else 'N/A',
            'DGR End Date': dgr.findtext('.//abr:endorsedTo', namespaces=namespaces, default='N/A').strip() if dgr else 'N/A',
            'DGR Item Number': dgr.findtext('.//abr:itemNumber', namespaces=namespaces, default='N/A').strip() if dgr else 'N/A',
            'DGR Funds': 'N/A'  # Default value as no funds element is shown in the example
        }

        charity_info = {
            'Charity Type': 'N/A',
            'Charity URL': 'N/A',
            'Charity Type Effective From': 'N/A',
            'Charity Type Effective To': 'N/A',
            'Endorsement Date': 'N/A',
            'Income Tax Exception': 'N/A',
            'GST Concession': 'N/A',
            'FBT Rebate': 'N/A',
            'FBT Exemption': 'N/A',
            'ACNC Registration Status': 'N/A',
            'ACNC Registration Effective From': 'N/A',
            'ACNC Registration Effective To': 'N/A'
        }

        abn = business_entity.findtext('.//abr:ABN/abr:identifierValue', namespaces=namespaces, default='N/A').strip()

        for charity_type in charity_types:
            charity_info['Charity Type'] = charity_type.findtext('.//abr:charityTypeDescription', namespaces=namespaces, default='N/A').strip()
            charity_info['Charity Type Effective From'] = charity_type.findtext('.//abr:effectiveFrom', namespaces=namespaces, default='N/A').strip()
            charity_info['Charity Type Effective To'] = charity_type.findtext('.//abr:effectiveTo', namespaces=namespaces, default='N/A').strip()
            charity_info['Charity URL'] = f'https://www.acnc.gov.au/charity/charities?search={abn}'

        for endorsement in endorsements:
            endorsement_type = endorsement.findtext('.//abr:endorsementType', namespaces=namespaces, default='N/A').strip()
            effective_from = endorsement.findtext('.//abr:effectiveFrom', namespaces=namespaces, default='N/A').strip()
            effective_to = endorsement.findtext('.//abr:effectiveTo', namespaces=namespaces, default='N/A').strip()
            if endorsement_type == 'GST Concession':
                charity_info['GST Concession'] = f'Yes (From: {effective_from}, To: {effective_to})'
            elif endorsement_type == 'Income Tax Exemption':
                charity_info['Income Tax Exception'] = f'Yes (From: {effective_from}, To: {effective_to})'
            elif endorsement_type == 'FBT Exemption':
                charity_info['FBT Exemption'] = f'Yes (From: {effective_from}, To: {effective_to})'
            elif endorsement_type == 'FBT Rebate':
                charity_info['FBT Rebate'] = f'Yes (From: {effective_from}, To: {effective_to})'

        if acnc_registration is not None:
            charity_info['ACNC Registration Status'] = acnc_registration.findtext('.//abr:status', namespaces=namespaces, default='N/A').strip()
            charity_info['ACNC Registration Effective From'] = acnc_registration.findtext('.//abr:effectiveFrom', namespaces=namespaces, default='N/A').strip()
            charity_info['ACNC Registration Effective To'] = acnc_registration.findtext('.//abr:effectiveTo', namespaces=namespaces, default='N/A').strip()

        abn_info = {
            'ABN': abn,
            'Organisation Name': business_entity.findtext('.//abr:mainName/abr:organisationName', namespaces=namespaces, default='N/A').strip(),
            'Entity Status': business_entity.findtext('.//abr:entityStatus/abr:entityStatusCode', namespaces=namespaces, default='N/A').strip(),
            'Effective From': business_entity.findtext('.//abr:entityStatus/abr:effectiveFrom', namespaces=namespaces, default='N/A').strip(),
            'State': business_entity.findtext('.//abr:mainBusinessPhysicalAddress/abr:stateCode', namespaces=namespaces, default='N/A').strip(),
            'Postcode': business_entity.findtext('.//abr:mainBusinessPhysicalAddress/abr:postcode', namespaces=namespaces, default='N/A').strip(),
            'GST Effective From': business_entity.findtext('.//abr:goodsAndServicesTax/abr:effectiveFrom', namespaces=namespaces, default='N/A').strip(),
            'Record Last Updated': business_entity.findtext('.//abr:recordLastUpdatedDate', namespaces=namespaces, default='N/A').strip(),
            'ASIC Number': business_entity.findtext('.//abr:ASICNumber', namespaces=namespaces, default='N/A').strip(),
            'Entity Type': business_entity.findtext('.//abr:entityType/abr:entityDescription', namespaces=namespaces, default='N/A').strip(),
            'Identifier Value': business_entity.findtext('.//abr:ABN/abr:identifierValue', namespaces=namespaces, default='N/A').strip(),
            'Date Retrieved': root.findtext('.//abr:dateTimeRetrieved', namespaces=namespaces, default='N/A').strip(),
            **dgr_info,
            **charity_info
        }
        
        return abn_info
    except ET.ParseError as e:
        print(f"XML parsing error: {e}")
        return None
    except ValueError as e:
        print(f"Date parsing error: {e}")
        return None

def display_single_abn_info(abn_info, abn):
    if abn_info:
        result_text = "\n".join([f"{key}: {value}" for key, value in abn_info.items()])
    else:
        result_text = f"Details for ABN {abn} could not be retrieved or are missing."
    
    single_abn_result.config(text=result_text)

def fetch_and_display_single_abn():
    abn = abn_entry.get().strip()
    api_key = api_key_entry.get().strip()
    if not abn:
        messagebox.showwarning("Warning", "Please enter an ABN.")
        return
    if not api_key:
        messagebox.showwarning("Warning", "Please enter an API key.")
        return
    abn_info = fetch_abn_details(abn, api_key)
    parsed_info = parse_abn_xml(abn_info) if abn_info else None
    display_single_abn_info(parsed_info, abn)

def create_excel_sheet(abns, api_key):
    df_list = []
    missing_abns = []
    date_retrieved = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Initialize progress bar
    progress_bar['maximum'] = len(abns)
    progress_bar['value'] = 0
    status_label.config(text=f"0/{len(abns)} ABNs processed")

    with ThreadPoolExecutor() as executor:
        results = list(executor.map(lambda abn: fetch_abn_details(abn, api_key), abns))
    for i, (abn, result) in enumerate(zip(abns, results)):
        abn_info = parse_abn_xml(result)
        if abn_info:
            abn_link = f'https://abr.business.gov.au/AbnHistory/View?id={abn_info["ABN"]}'
            asic_number = abn_info["ASIC Number"]
            asic_link = f'https://connectonline.asic.gov.au/RegistrySearch/faces/landing/panelSearch.jspx?searchType=OrgAndBusNm&searchText={asic_number}'
            df_list.append([
                abn_link, abn_info['Organisation Name'], abn_info['Entity Type'],
                abn_info['GST Effective From'], abn_info['Entity Status'], abn_info['State'],
                abn_info['Postcode'], asic_link, abn_info['Record Last Updated'],
                abn_info['GST Effective From'], abn_info['Effective From'], date_retrieved,
                abn_info['DGR Endorsement'], abn_info['DGR Start Date'], abn_info['DGR End Date'],
                abn_info['DGR Item Number'], abn_info['DGR Funds'], abn_info['Charity Type'],
                abn_info['Charity URL'], abn_info['Charity Type Effective From'], abn_info['Charity Type Effective To'], 
                abn_info['Endorsement Date'], abn_info['Income Tax Exception'], abn_info['GST Concession'],
                abn_info['FBT Rebate'], abn_info['FBT Exemption'], abn_info['ACNC Registration Status'],
                abn_info['ACNC Registration Effective From'], abn_info['ACNC Registration Effective To']
            ])
        else:
            missing_abns.append(abn)
        
        # Update progress bar and status label
        progress_bar['value'] += 1
        status_label.config(text=f"{i + 1}/{len(abns)} ABNs processed")
        root.update_idletasks()

    if df_list:
        file_path = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")])
        if file_path:
            df = pd.DataFrame(df_list, columns=[
                'ABN', 'Organisation Name', 'Entity Type', 'GST Registration', 'Status',
                'State', 'Postcode', 'ASIC Number', 'Record Last Updated', 'GST Effective From',
                'Effective From', 'Date Retrieved', 'DGR Endorsement', 'DGR Start Date',
                'DGR End Date', 'DGR Item Number', 'DGR Funds', 'Charity Type',
                'Charity URL', 'Charity Type Effective From', 'Charity Type Effective To', 'Endorsement Date',
                'Income Tax Exception', 'GST Concession', 'FBT Rebate', 'FBT Exemption',
                'ACNC Registration Status', 'ACNC Registration Effective From', 'ACNC Registration Effective To'
            ])
            writer = pd.ExcelWriter(file_path, engine='xlsxwriter')
            df.to_excel(writer, index=False, sheet_name='ABN Details')

            # Get the xlsxwriter workbook and worksheet objects.
            workbook = writer.book
            worksheet = writer.sheets['ABN Details']

            # Iterate over the DataFrame to set the URL for the ABN, ASIC and Charity columns.
            for idx, link in enumerate(df['ABN']):
                abn_id = link.split('=')[-1]
                worksheet.write_url(idx + 1, 0, link, string=abn_id)
            for idx, link in enumerate(df['ASIC Number']):
                asic_id = link.split('=')[-1]
                worksheet.write_url(idx + 1, 7, link, string=asic_id)
            for idx, link in enumerate(df['Charity URL']):
                if link != 'N/A':
                    worksheet.write_url(idx + 1, 18, link, string='Charity URL')

            # Set column widths
            column_widths = [
                12, 32.14, 28.86, 15.71, 10.57, 5.57, 7.86, 11.71, 18.57, 17, 13.43, 17.86, 
                16.71, 13.86, 12.43, 16.86, 10.14, 11.86, 11.29, 25.43, 22.14, 16.71, 34, 34, 
                12.86, 33.43, 23.43, 31.29, 28.71
            ]

            for i, width in enumerate(column_widths):
                worksheet.set_column(i, i, width)

            writer.close()
            
            message = f"Excel file has been created successfully at {file_path}."
            if missing_abns:
                missing_df = pd.DataFrame(missing_abns, columns=['ABN'])
                missing_path = file_path.replace(".xlsx", "_missing.xlsx")
                missing_df.to_excel(missing_path, index=False)
                message += f" Additionally, 'missing_abns.xlsx' has been created with {len(missing_abns)} missing ABN(s) at {missing_path}."
            messagebox.showinfo("Success", message)
        else:
            messagebox.showwarning("Warning", "File save operation was cancelled.")
    else:
        messagebox.showinfo("Information", "No valid ABN details found to save.")

def fetch_and_create_excel():
    abn_input = fetch_text.get("1.0", "end-1c").strip()
    api_key = api_key_entry.get().strip()
    if not abn_input:
        messagebox.showwarning("Warning", "Please enter ABN(s) in the text box.")
        return
    if not api_key:
        messagebox.showwarning("Warning", "Please enter an API key.")
        return
    abns = [abn.strip() for abn in abn_input.splitlines() if abn.strip()]
    create_excel_sheet(abns, api_key)

root = tk.Tk()
root.title("ABN Details Fetcher")

main_frame = ttk.Frame(root, padding="10")
main_frame.pack(fill=tk.BOTH, expand=True)

ttk.Label(main_frame, text="Enter API Key:").grid(column=0, row=0, padx=10, pady=10)
api_key_entry = ttk.Entry(main_frame, width=30)
api_key_entry.insert(0, api_key if api_key else "")
api_key_entry.grid(column=1, row=0, padx=10, pady=10)

ttk.Label(main_frame, text="Enter ABN:").grid(column=0, row=1, padx=10, pady=10)
abn_entry = ttk.Entry(main_frame, width=30)
abn_entry.grid(column=1, row=1, padx=10, pady=10)
ttk.Button(main_frame, text="Fetch ABN Info", command=fetch_and_display_single_abn).grid(column=2, row=1, padx=10, pady=10)

single_abn_result = ttk.Label(main_frame, text="", justify=tk.LEFT)
single_abn_result.grid(column=0, row=2, columnspan=3, padx=10, pady=10, sticky='w')

ttk.Label(main_frame, text="Enter ABNs (one per line for batch fetch):").grid(column=0, row=3, columnspan=3, padx=10, pady=10, sticky='w')
fetch_text = tk.Text(main_frame, height=10, width=50)
fetch_text.grid(column=0, row=4, columnspan=3, padx=10, pady=10)
ttk.Button(main_frame, text="Fetch & Create Excel", command=fetch_and_create_excel).grid(column=1, row=5, padx=10, pady=10)

# Add progress bar and status label
progress_bar = ttk.Progressbar(main_frame, orient='horizontal', mode='determinate')
progress_bar.grid(column=0, row=6, columnspan=3, padx=10, pady=10, sticky='we')

status_label = ttk.Label(main_frame, text="")
status_label.grid(column=0, row=7, columnspan=3, padx=10, pady=10)

# Add "Created by" section
created_by_frame = ttk.Frame(root)
created_by_frame.pack(fill=tk.X, side=tk.BOTTOM)

created_by = ttk.Label(created_by_frame, text="Created by InfinityHack3r", foreground="blue", cursor="hand2")
created_by.pack(padx=10, pady=10)
created_by.bind("<Button-1>", lambda e: webbrowser.open_new("https://github.com/InfinityHack3r"))

root.mainloop()
