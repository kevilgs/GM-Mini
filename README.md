# GM MINI offline workbook generator

1. Keep `GM MINI- AUGUST 2026.xlsx` in this folder. It is the untouched visual template.
2. Double-click `run_generator.bat`.
3. Open `http://127.0.0.1:5050` if the browser does not open automatically.
4. Enter the period, targets, opening balances, and prior-year daily values.
5. Click **Generate offline workbook**. The output is saved in `generated_workbooks`.

The generated workbook needs no network connection. Employees enter their daily values directly in Excel and Excel calculates the included formulas locally.

The first implementation preserves the GM MINI daily and LINK layouts. It covers the core commodity, division, rake-stock, target, MTD, FYTD, and prior-year calculation paths. The detailed daily port/silo/container panels are preserved as formatted Excel areas and are left for the employee to enter.
