import pandas as pd
from statsmodels.tsa.holtwinters import SimpleExpSmoothing
from sklearn.metrics import mean_absolute_error, mean_squared_error

# Load and clean
df = pd.read_excel("Expenses.xlsx", sheet_name="Table 1")
df = df.dropna(subset=['Unnamed: 1']).copy()
df.columns = ['Code', 'Expense Type', '2024', '2023', '2022', '2021', '2020', '2019']
df = df.drop(columns=['Code'])
df = df[['Expense Type', '2019', '2020', '2021', '2022', '2023', '2024']]

# Melt to long format
df_melted = df.melt(id_vars='Expense Type', var_name='Year', value_name='Amount')
df_melted['Year'] = df_melted['Year'].astype(int)

future_years = [2025, 2026, 2027]
train_years = [2019, 2020, 2021, 2022, 2023, 2024]
all_years = train_years + future_years

final_data = []

for expense in df['Expense Type']:
    sub_df = df_melted[df_melted['Expense Type'] == expense].sort_values(by='Year')
    if len(sub_df) >= 3:
        model = SimpleExpSmoothing(sub_df['Amount'], initialization_method="estimated")
        fit = model.fit()
        forecast_values = fit.predict(start=0, end=len(all_years)-1)
        forecast_series = pd.Series(forecast_values.values, index=all_years).round(2)

        # Get actuals
        actuals = sub_df.set_index('Year')['Amount'].round(2).to_dict()

        row = {'Expense Type': expense}
        for year in all_years:
            if year in actuals:
                row[year] = actuals[year]
            row[f"{year} (Pred)"] = forecast_series[year]

        mae = mean_absolute_error(sub_df['Amount'], forecast_series[train_years])
        mse = mean_squared_error(sub_df['Amount'], forecast_series[train_years])
        row['MAE'] = round(mae, 2)
        row['MSE'] = round(mse, 2)

        final_data.append(row)

# Convert to DataFrame
result_df = pd.DataFrame(final_data)

# Add total row
total_row = {'Expense Type': 'TOTAL FORECAST'}
for year in all_years:
    pred_col = f"{year} (Pred)"
    if pred_col in result_df.columns:
        total_row[pred_col] = result_df[pred_col].sum().round(2)
result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)

# Save to Excel
with pd.ExcelWriter("Exponential_Smoothing_Forecast_2019_to_2027.xlsx", engine="xlsxwriter") as writer:
    result_df.to_excel(writer, sheet_name="Exponential_Smoothing_Forecast", index=False)

    workbook = writer.book
    worksheet = writer.sheets["Exponential_Smoothing_Forecast"]

    num_format = workbook.add_format({'num_format': '#,##0.00'})
    highlight_future = workbook.add_format({'bg_color': '#FFF2CC', 'num_format': '#,##0.00'})
    highlight_pred = workbook.add_format({'bg_color': '#D9EAD3', 'num_format': '#,##0.00'})

    # Apply number format
    worksheet.set_column(1, len(result_df.columns)-1, 15, num_format)

    for i, col_name in enumerate(result_df.columns[1:], start=1):
        if "(Pred)" in str(col_name):
            worksheet.set_column(i, i, 15, highlight_pred)
        elif col_name in [2025, 2026, 2027]:
            worksheet.set_column(i, i, 15, highlight_future)
        else:
            worksheet.set_column(i, i, 15, num_format)
            
print("✅ Exponential Smoothing Forecast completed and exported.")
