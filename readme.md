<p align="center">

<h1 align="center">ABN Bulk Lookup</h1>
<h2 align="center">🐍 Python | 🧾 ABN Details Fetcher | 📊 Data Processing | 📝 Excel Export</h2>
</p>

## 🌟 **Introduction**

ABN Bulk Lookup is a Python application that fetches and processes Australian Business Numbers (ABNs) details from the ABN API. The application allows users to fetch details for individual ABNs or batch process multiple ABNs and save the details to an Excel file.

**Project name:** [ABN Bulk Lookup](https://github.com/InfinityHack3r/abnBulkLookup)

**Language:** [Python](https://www.python.org/)

**Libraries:**
- [Requests](https://docs.python-requests.org/en/master/)
- [Pandas](https://pandas.pydata.org/)
- [Openpyxl](https://openpyxl.readthedocs.io/en/stable/)
- [Python-dotenv](https://saurabh-kumar.com/python-dotenv/)
- [XlsxWriter](https://xlsxwriter.readthedocs.io/)
- [PyInstaller](https://www.pyinstaller.org/) (optional)

## 🚀 **Features**

- **Fetched ABN Details:**
    - 🔢 ABN
    - 🏢 Organisation Name
    - 🏷️ Entity Type
    - 🧾 GST Registration
    - 📌 Status
    - 🌐 State
    - 🏷️ Postcode
    - 🔖 ASIC Number
    - 🕒 Record Last Updated
    - 📅 GST Effective From
    - 📅 Effective From
    - 📅 Date Retrieved
    - 🏅 DGR Endorsement
    - 📅 DGR Start Date
    - 📅 DGR End Date
    - #️⃣ DGR Item Number
    - 💰 DGR Funds
    - 🏷️ Charity Type
    - 🌐 Charity URL
    - 📅 Charity Type Effective From
    - 📅 Charity Type Effective To
    - 📅 Endorsement Date
    - 🏷️ Income Tax Exception
    - 🧾 GST Concession
    - 💸 FBT Rebate
    - 💸 FBT Exemption
    - 🔖 ACNC Registration Status
    - 📅 ACNC Registration Effective From
    - 📅 ACNC Registration Effective To
    
- 📝 Display details for a single ABN.
- 📊 Batch process multiple ABNs and save the details to an Excel file.
- 💾 Save the generated Excel file to a custom location.
- 🔒 Load with or without a `.env` file containing the ABN API key.

## 🛠️ **Requirements**

- Python `3.11.5` or higher
- ABN API key (get one [here](https://abr.business.gov.au/Tools/WebServices))
- Required Python packages found in `requirements.txt`

## 📖 **Usage**

1. Clone the repository:

```sh
git clone https://github.com/InfinityHack3r/abnBulkLookup.git
cd abnBulkLookup
```

2. You can install the required packages using:

```sh
pip install -r requirements.txt
```

3. (Optional) Create a `.env` file in the project directory and add your ABN API key:

```
ABN_API_KEY=your_api_key_here
```

4. Run the application:

```sh
python abnBulkLookup.py
```

5. Enter your ABN API key and ABN(s) in the provided fields.

6. Click "Fetch ABN Info" to fetch details for a single ABN or "Fetch & Create Excel" to batch process multiple ABNs and save the details to an Excel file.

## 🛠️ **Building the Executable**

You can build the executable using PyInstaller manually.

### Manually

1. Install PyInstaller:

```sh
pip install pyinstaller
```

2. Build the executable:

```sh
pyinstaller --onefile --noconsole abnBulkLookup.py --hidden-import requests --hidden-import pandas --hidden-import openpyxl --hidden-import dotenv --hidden-import xlsxwriter
```

3. The executable will be found in the `dist` directory.

## 📜 **License**

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 👤 **Created by**

This application was created by [InfinityHack3r](https://github.com/InfinityHack3r).
