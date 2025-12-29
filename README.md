# Expense Forecasting Methods

This repository contains forecasting methods implemented in Python, inspired by tasks encountered during my internship experience.  
All data used is anonymized or simulated for confidentiality.

## Project Overview

During my internship, I worked on forecasting techniques to project future expenses based on historical data.  
This repository demonstrates several basic forecasting approaches implemented in Python:

- Linear Regression
- Moving Average
- Exponential Smoothing

The focus is on method implementation and result interpretation, not on specific company data.

## Methods Included

### Linear Regression
The `linear_regression.py` script demonstrates a simple linear regression model applied to historical expense data
to forecast future trends. It includes:
- Data input and validation
- Model training
- Forecast generation
- Error evaluation

### Moving Average
The `moving_average.py` script implements a moving average forecast, suitable for smoothing time series data
without trend or seasonality.

### Exponential Smoothing
The `exponential_smoothing.py` script applies single exponential smoothing using `statsmodels`
to capture level trends in expense series.

## How to Use

1. Prepare your data in a CSV file with a consistent date/time index and expense value column.
2. Modify the script to point to your CSV dataset path.
3. Run the Python script to generate forecast results and evaluation metrics.

```bash
python linear_regression.py
