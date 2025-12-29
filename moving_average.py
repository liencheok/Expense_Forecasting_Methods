import pandas as pd
from sklearn.metrics import mean_absolute_error, mean_squared_error

# Load and clean data
df = pd.read_excel("Expenses.xlsx", sheet_name="Table 1")
df = df.dropna(subset=['Unnamed: 1']).copy()
df.columns = ['Code', 'Expense Type', '2024', '2023', '2022', '2021', '2020', '2019']
df = df.drop(columns=['Code'])
df = df[['Expense Type', '2019', '2020', '2021', '2022', '2023', '2024']]

# Melt to long format
df_melted = df.melt(id_vars='Expense Type', var_name='Year', value_name='Amount')
df_melted['Year'] = df_melted['Year'].astype(int)

train_years = [2019, 2020, 2021, 2022, 2023, 2024]
future_years = [2025, 2026, 2027]

final_data = []

for expense in df['Expense Type']:
    sub_df = df_melted[df_melted['Expense Type'] == expense].sort_values('Year')
    values = sub_df['Amount'].tolist()
    years = sub_df['Year'].tolist()

    if len(values) >= 3:
        full_years = years.copy()
        forecast_train = []
        for i in range(len(values)):
            if i < 3:
                forecast_train.append(None)
            else:
                avg = sum(values[i-3:i]) / 3
                forecast_train.append(round(avg, 2))

        # MAE and MSE
        preds_train_clean = [p for p in forecast_train[3:] if p is not None]
        actuals_train_clean = values[3:]
        mae = mean_absolute_error(actuals_train_clean, preds_train_clean)
        mse = mean_squared_error(actuals_train_clean, preds_train_clean)

        # Forecast future years
        forecast_future = values.copy()
        for _ in future_years:
            if len(forecast_future) >= 3:
                avg = sum(forecast_future[-3:]) / 3
                forecast_future.append(round(avg, 2))

        row = {'Expense Type': expense}
        for y, val in zip(train_years, values):
            row[y] = val

        # Training predictions
        for i, y in enumerate(train_years):
            pred_col = f"{y} (Pred)"
            row[pred_col] = forecast_train[i] if forecast_train[i] is not None else ""

        # Future forecasts
        row[2025] = forecast_future[-3]
        row[2026] = forecast_future[-2]
        row[2027] = forecast_future[-1]

        row["MAE"] = round(mae, 2)
        row["MSE"] = round(mse, 2)
        final_data.append(row)

# Create DataFrame
result_df = pd.DataFrame(final_data)

# Organize columns
cols = ['Expense Type']
for y in train_years:
    cols += [y, f"{y} (Pred)"]
cols += [2025, 2026, 2027, "MAE", "MSE"]
result_df = result_df[cols]

# Add TOTAL row
total_row = {'Expense Type': 'TOTAL FORECAST'}
for col in result_df.columns[1:]:
    if result_df[col].dtype in ['float64', 'int64']:
        total_row[col] = result_df[col].sum().round(2)
result_df = pd.concat([result_df, pd.DataFrame([total_row])], ignore_index=True)

# Export to Excel
with pd.ExcelWriter("Moving_Average_Forecast_2019_to_2027.xlsx", engine="xlsxwriter") as writer:
    result_df.to_excel(writer, sheet_name="Moving_Average_Forecast", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Moving_Average_Forecast"]

    # Format
    num_format = workbook.add_format({'num_format': '#,##0.00'})
    highlight = workbook.add_format({'bg_color': '#FFF2CC', 'num_format': '#,##0.00'})

    worksheet.set_column(1, len(result_df.columns)-1, 15, num_format)
    for col_name in [2025, 2026, 2027] + [f"{y} (Pred)" for y in train_years if y >= 2022]:
        if col_name in result_df.columns:
            col_idx = result_df.columns.get_loc(col_name)
            worksheet.set_column(col_idx, col_idx, 15, highlight)

print("✅ Moving Average Forecast completed and exported.")
