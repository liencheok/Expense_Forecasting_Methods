import pandas as pd
from sklearn.linear_model import LinearRegression
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

# Prepare years
train_years = list(range(2019, 2025))
future_years = [2025, 2026, 2027]
all_years = train_years + future_years

final_rows = []

for expense in df['Expense Type'].unique():
    row = {'Expense Type': expense}
    sub_df = df_melted[df_melted['Expense Type'] == expense]
    
    if len(sub_df) >= 3:
        model = LinearRegression().fit(sub_df[['Year']], sub_df['Amount'])
        predictions = model.predict(pd.DataFrame({'Year': all_years}))
        predictions = [max(0, round(p, 2)) for p in predictions]

        actuals, preds_for_error = [], []

        for i, year in enumerate(train_years):
            actual_val = sub_df[sub_df['Year'] == year]['Amount'].values
            actual = round(actual_val[0], 2) if len(actual_val) > 0 else None
            pred = predictions[i]
            row[str(year)] = actual
            row[f"{year} (Pred)"] = pred
            if actual is not None:
                actuals.append(actual)
                preds_for_error.append(pred)

        for i, year in enumerate(future_years):
            row[f"{year} (Pred)"] = predictions[len(train_years) + i]

        if len(actuals) >= 2:
            row["MAE"] = round(mean_absolute_error(actuals, preds_for_error), 2)
            row["MSE"] = round(mean_squared_error(actuals, preds_for_error), 2)
        else:
            row["MAE"] = None
            row["MSE"] = None

        final_rows.append(row)

# Output to Excel
df_final = pd.DataFrame(final_rows)
with pd.ExcelWriter("Linear_Regression_Forecast_2019_to_2027.xlsx", engine="xlsxwriter") as writer:
    df_final.to_excel(writer, sheet_name="Linear_Regression_Forecast", index=False)
    workbook = writer.book
    worksheet = writer.sheets["Linear_Regression_Forecast"]
    num_format = workbook.add_format({'num_format': '#,##0.00'})
    highlight = workbook.add_format({'bg_color': '#FFF2CC', 'num_format': '#,##0.00'})
    worksheet.set_column(1, len(df_final.columns)-1, 15, num_format)
    for idx, col_name in enumerate(df_final.columns):
        if "(Pred)" in str(col_name):
            worksheet.set_column(idx, idx, 15, highlight)

print("✅ Linear Regression Forecast completed and exported.")
