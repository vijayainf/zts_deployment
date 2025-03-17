# ZTS Deployment JSON Generator

This repository contains a tool for generating a ZTS deployment JSON file by merging data from an Excel file with a JSON template. It was designed to support multiple SITE column values, so you can generate one output per SITE from the same Excel file.

## Features

- **Excel DP Data Processing:**  
  Reads data from the "DP" sheet in an Excel file. The "NE Parameter Name" column is used for keys, and a specified SITE column provides the corresponding values.
  
- **Multiple SITE Support:**  
  You can specify multiple SITE column names (comma-separated) to generate separate output files for each SITE. Output files are named as `<SITE>_zts_24.7_mp1.json`, along with corresponding log (`<SITE>_zts_24.7_mp1.log`) and variable (`<SITE>_zts_24.7_mp1.var`) files.

- **Robust Parameter Updating:**  
  The tool performs case-insensitive and punctuation-insensitive matching of keys (e.g., "vnfname", "imageRegistry", "clustername") so that values are correctly updated from the Excel sheet.

- **Composite Section Merging:**  
  For composite sections (e.g., "UmIdpConfig" under "chartValues"), the tool merges related keys from the Excel file into the corresponding JSON structure.

## Requirements

- Python 3.x
- [openpyxl](https://pypi.org/project/openpyxl/)

Install the required package using:
```bash
pip install -r requirements.txt
```

## Usage

1. **Prepare Your Files:**  
   - Ensure your Excel file contains a "DP" sheet with a column named **NE Parameter Name** and one or more SITE columns.
   - Have your JSON template file (e.g., `zts_template.json`) ready.

2. **Run the Generator Script:**
   ```bash
   python e2503_orb_zts_deployment_generator_json.py
   ```
3. **Provide the Required Inputs:**  
   - **Excel file path:** Path to the Excel file.
   - **JSON template file path:** Path to your JSON template file.
   - **SITE columns:** Comma-separated list of SITE column names as they appear in the Excel file.
   - **Output directory:** Directory where the output files will be saved.

4. **Output Files:**  
   For each SITE value specified, the tool generates three files in the output directory:
   - `<SITE>_zts_24.7_mp1.json` – The updated JSON configuration.
   - `<SITE>_zts_24.7_mp1.log` – Log file with processing details.
   - `<SITE>_zts_24.7_mp1.var` – A variable file containing the raw DP data in JSON format.

## Repository Structure

- `e2503_orb_zts_deployment_generator_json.py` – The main Python script.
- `README.md` – This README file.
- `requirements.txt` – List of required Python packages.

## Author

- **Vijayakumar R**
- **CFX-5000 Team**
- **Email:** vijaya.r.ext@nokia.com

## Notes

- Ensure that the "DP" sheet in your Excel file has the correct headers.
- The tool performs normalization (via lower-casing and removal of non-alphanumeric characters) to match parameter names between the Excel file and JSON template.
- If you encounter any issues or discrepancies in the output, refer to the log files generated for each SITE for debugging details.

Enjoy using the ZTS Deployment JSON Generator!