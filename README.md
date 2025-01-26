# Excel-Automation-Tool

## Overview
The **Excel-Automation-Tool** is a Python-based solution designed to streamline Excel data processing and automate repetitive tasks. This tool is ideal for businesses looking to improve efficiency by filtering data, generating reports, and integrating with cloud platforms like Google Sheets.

**Documentation Available in Multiple Languages**
- [English README.md](README.md) - This document.
- [日本語 README_ja.md](README_ja.md) - 日本語ドキュメントはこちら。

## Features
- **User Input Support**: Allows users to specify an Excel file for processing.
- **Data Filtering**: Automatically filters data based on sales thresholds.
- **Graph Generation**: Creates bar charts visualizing sales data by category.
- **PDF Report Generation**: Exports a summarized sales report in PDF format.
- **Google Sheets Integration**: Uploads processed data to Google Sheets for easy access.

## Project Structure
The project follows this directory structure:
```plaintext
Excel-Automation-Tool/
│── sample_data/         # Folder for sample Excel files and output files
│   ├── sample.xlsx      # Example Excel file (to be provided by the user)
│   ├── high_sales.xlsx  # Output file for filtered sales data
│   ├── sales_chart.png  # Generated sales graph
│   ├── sales_report.pdf # Generated PDF report
│── venv/                # Virtual environment (not included in GitHub)
│── main.py              # Main Python script
│── requirements.txt     # Dependencies file
│── README.md            # Documentation file
│── README_ja.md         # Documentation file in Japanese (日本語のドキュメント)
│── credentials.json     # Google Sheets API credentials (not included in GitHub)
```

## Requirements
Ensure you have the necessary dependencies installed before running the tool:
```bash
pip install pandas openpyxl matplotlib fpdf gspread oauth2client
```

## Installation & Usage
1. Clone the repository:
```bash
git clone https://github.com/JourneySculptor/Excel-Automation-Tool.git
cd Excel-Automation-Tool
```
2. Install dependencies:
```bash
pip install -r requirements.txt
```
3. Place your Excel file in the `sample_data/` folder.
4. Run the script:
```bash
python main.py
```
5. Follow the prompts to process the Excel file.

## Example Output
### Filtered Data:
The script extracts rows where sales exceed 10,000 and saves them in `high_sales.xlsx`.

### Generated Graph:
A bar chart showing sales by category is saved as `sales_chart.png`.

### PDF Report:
A summary report is generated as `sales_report.pdf`.

### Google Sheets Upload:
If configured correctly, the filtered data is uploaded to Google Sheets.

### Google Sheets Setup
To enable Google Sheets integration:
1. Create a **Google Cloud project** and enable the **Google Sheets API**. 👉 [Google Cloud Console](https://console.cloud.google.com/)
2. Generate a **Service Account Key** (`credentials.json`).
3. Place `credentials.json` in the **project root directory** (`Excel-Automation-Tool/`).
4. Run the script to automatically upload data to Google Sheets.

> **Do NOT upload `credentials.json` to GitHub!** Add it to `.gitignore` to keep it secure.

## Future Enhancements
- Implement GUI for non-technical users.
- Add email automation for sending reports.
- Enhance error handling for better user experience.

💡 **Got ideas or found a bug?**  
Feel free to open an issue on [GitHub Issues](https://github.com/JourneySculptor/Excel-Automation-Tool/issues) to share your feedback!

## License
This project is licensed under the MIT License.

## Contact
For inquiries or collaboration opportunities, feel free to connect with me:
- [LinkedIn](https://www.linkedin.com/in/yuka-yamaguchi-214290342)
- [GitHub Profile](https://github.com/JourneySculptor)