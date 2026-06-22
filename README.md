# Yahoo Finance Historical Data Scraper

A lightweight, automated Python tool designed to research a company's public trading history and extract daily historical stock price data directly from Yahoo Finance into a styled Excel workbook.

## Features
- **Automated Research:** Dynamically identifies a company's initial public trading date (Inception/IPO Date).
- **Custom Timeframe:** Pulls complete daily historical data from the inception date up to December 31, 2025.
- **Clean Excel Formatting:** Automatically exports data into a structured spreadsheet with auto-fitted columns and standard financial metrics.

## Extracted Data Fields
- Date (YYYY-MM-DD)
- Open / High / Low / Close
- Adjusted Close
- Volume

## Output Preview
<img width="1364" height="702" alt="image" src="https://github.com/user-attachments/assets/da969f61-f506-4835-9f9d-24be81042123" />


## Setup & Installation

1. Clone or download this repository to your local machine.
2. Install the required Python dependencies:
   ```bash
   pip install yfinance pandas openpyxl

