import yfinance as yf
import pandas as pd
from datetime import datetime

def get_historical_data(ticker_symbol):
    print(f"🕵️ Target Ticker: {ticker_symbol.upper()}")
    
    # 1. Initialize the ticker object
    ticker = yf.Ticker(ticker_symbol)
    
    # 2. Automatically retrieve the earliest available historical data date
    print("🔍 Researching company public inception date...")
    full_history = ticker.history(period="max")
    
    if full_history.empty:
        print("❌ Error: No data found. Please check the ticker symbol.")
        return
        
    inception_date = full_history.index.min().strftime('%Y-%m-%d')
    print(f"📅 Public Trading Start Date Found: {inception_date}")
    
    # 3. Filter data from inception up to December 31, 2025
    end_date = "2026-01-31"
    print(f"📥 Extracting historical price data up to {end_date}...")
    
    # yfinance dates are exclusive on the end date, so we use 2026-01-01 to capture all of Dec 2025
    filtered_data = ticker.history(start=inception_date, end="2026-01-01")
    
    # 4. Clean up and format the dataset
    filtered_data = filtered_data.reset_index()
    
    # Keep only the requested standard columns
    columns_to_keep = ['Date', 'Open', 'High', 'Low', 'Close', 'Volume']
    # Check if 'Adj Close' exists (yfinance names it 'Capital Gains' or folds it into Close depending on asset class)
    if 'Adj Close' in filtered_data.columns:
        columns_to_keep.insert(5, 'Adj Close')
    else:
        # For standard stock tickers, 'Close' is already split-adjusted; we can duplicate or calculate it if needed
        filtered_data['Adj Close'] = filtered_data['Close']
        columns_to_keep.insert(5, 'Adj Close')
        
    final_df = filtered_data[columns_to_keep]
    
    # Ensure Date column does not carry timezone data for clean Excel export
    final_df['Date'] = final_df['Date'].dt.tz_localize(None)
    
    # 5. Export to a styled Excel File
    file_name = f"{ticker_symbol.upper()}_historical_data.xlsx"
    print(f"📊 Generating styled Excel file: {file_name}...")
    
    with pd.ExcelWriter(file_name, engine='openpyxl') as writer:
        final_df.to_excel(writer, sheet_name="Historical Price Data", index=False)
        
        # Access openpyxl features for auto-adjusting column widths
        workbook = writer.book
        worksheet = writer.sheets["Historical Price Data"]
        worksheet.views.sheetView[0].showGridLines = True
        
        for col in worksheet.columns:
            max_len = max(len(str(cell.value or '')) for cell in col)
            col_letter = col[0].column_letter
            worksheet.column_dimensions[col_letter].width = max(max_len + 3, 12)
            
    print(f"✅ Success! Your file is ready. Total rows extracted: {len(final_df)}")

if __name__ == "__main__":
    # Input any ticker here (e.g., 'AAPL', 'MSFT', or Philippine stocks like 'ALI.PS')
    target_ticker = input("Enter company ticker symbol (e.g., AAPL): ").strip()
    get_historical_data(target_ticker)
