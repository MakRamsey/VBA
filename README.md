# Visual Basics for Applications (VBA)

Greetings,

Welcome to the VBA-Scripting repository! Within this project, VBA Scripting was used to both analyze and aggregate generated stock market data into indivudal annual stock Ticker summary statistics/metrics ('Ticker Name', 'Yearly Change', 'Percentage Change' and 'Total Stock Volume'). The raw data included daily reported values corresponding to a specific stock's 'date', 'open', 'high', 'low', 'close' and 'vol' (volume). The initial input .xlsx file contained this raw data across three separate sheets, each corresponding to a different year (2018, 2019 & 2020). The VBA script reads the incoming .xlsx file and subsequenlty generates the desired aggregates for every stock ticker included across all three sheets. Global metrics for each sheet were calculated in order to identify which stock represents the 'Greatest % Increase', 'Greatest % Decrease' and 'Greatest Total Volume'. Conditional formatting was also applied to highlight positive delta as green and a negative delta as red.


**Repository Structure:**

- 'Resource' directory: Contains raw generated stock market data ("Multiple_year_stock_data.xlsx")
- 'Result_Output' directory: Includes three .png files that depict the VBA script's results upon execution for each sheet
- 'VBA_Scripting.vbs': Executed VBA Scripting

Thank you!
