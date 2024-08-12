# Stock Data Processing VBA Script

## Overview

This VBA script, `ProcessStocks`, processes stock data across multiple worksheets in an Excel workbook. It calculates key metrics such as quarterly change, percent change, and total stock volume for each stock ticker. Additionally, it identifies the tickers with the greatest percentage increase, the greatest percentage decrease, and the greatest total volume. The results are displayed directly on each worksheet, with conditional formatting applied for easy visualization.

## How It Works

### Code Explanation

1. **Variables Initialization**:
    - The script begins by declaring variables to store the worksheet (`ws`), ticker symbol (`ticker`), stock prices (`openPrice`, `closePrice`), and total volume (`totalVolume`).
    - It also declares variables to track the row numbers and the greatest values for percentage increase, decrease, and volume.

2. **Loop Through Worksheets**:
    - The script loops through each worksheet in the workbook using `For Each ws In ThisWorkbook.Worksheets`.
    - For each worksheet, it identifies the last row of data and initializes a `summaryRow` to track where the results will be recorded.

3. **Header Setup**:
    - Column headers are set up in columns H to K for "Ticker," "Quarterly Change," "Percent Change," and "Total Stock Volume".

4. **Processing Each Ticker**:
    - The script enters a loop (`While i <= lastRow`) to process each ticker.
    - For each ticker:
        - It retrieves the ticker symbol and the opening price.
        - It calculates the total volume traded by summing the values in column G.
        - The closing price is taken from the last entry of the current ticker.
        - The quarterly change (difference between closing and opening prices) and percent change are calculated.
        - These values are recorded in the summary section starting at `summaryRow`.

5. **Identifying Greatest Values**:
    - As it processes each ticker, the script compares the current ticker's percent change and volume with the greatest values recorded so far. It updates the greatest values accordingly.

6. **Output Greatest Values**:
    - After processing all tickers, the script outputs the greatest percentage increase, decrease, and total volume in columns M to O of the worksheet.

7. **Conditional Formatting**:
    - Conditional formatting is applied to highlight positive changes in green and negative changes in red. This is done for both the quarterly change and percent change columns.

### How to Use the Script

1. **Add the Script to Your Workbook**:
   - Open Excel and press `Alt + F11` (Windows) or `Option + F11` (Mac) to open the VBA editor.
   - Insert a new module (`Insert > Module`) and paste the script into the module.

2. **Run the Script**:
   - Press `F5` within the VBA editor to run the script, or close the VBA editor and run the macro from the Excel interface (`Alt + F8`).

3. **View Results**:
   - The results are displayed starting in column H of each worksheet. The greatest values are shown in columns M to O.

### Example of Script Output

Here’s an example of what the script outputs in the worksheet:

| Ticker | Quarterly Change | Percent Change | Total Stock Volume |
|--------|------------------|----------------|--------------------|
| AAPL   |  15.00           | 10.50%         |  2,000,000         |
| MSFT   | -5.00            | -2.30%         |  1,500,000         |

- **Greatest % Increase**: AAPL (10.50%)
- **Greatest % Decrease**: MSFT (-2.30%)
- **Greatest Total Volume**: AAPL (2,000,000)

### Customization

Feel free to modify the script to:
- Change the columns where data is output.
- Add additional calculations.
- Customize the formatting.

### Requirements

- **Excel**: The script is compatible with both Windows and Mac versions of Excel.

### License

This script is open-source and available for modification and distribution. Feel free to use it in your projects.

