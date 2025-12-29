# Expense Forecasting Methods

This repository contains Python implementations of basic expense forecasting methods, inspired by tasks encountered during my internship.  
All example data referenced in this repository is anonymized or simulated to maintain confidentiality.

## Project Overview

During my internship, I worked with forecasting techniques to project future expenses based on historical data.  
This repository demonstrates several fundamental forecasting approaches in Python, including:

- Linear Regression
- Moving Average
- Exponential Smoothing

These methods are widely used in time series forecasting and provide a foundation for more advanced predictive analytics.

## Methods Included

### 🧮 Linear Regression
The `linear_regression.py` script demonstrates a simple linear regression model used to identify trends and forecast future values from historical data.

### 📊 Moving Average
The `moving_average.py` script implements a moving average approach that smooths out short-term fluctuations and highlights long-term trends in a series.

### 🔁 Exponential Smoothing
The `exponential_smoothing.py` script applies single exponential smoothing using the `statsmodels` library to account for underlying trends in the data.

## Usage

1. Prepare your own dataset in a CSV file with consistent chronological structure.
2. Modify the script to reference your dataset file path.
3. Run the script using Python:

```bash
python linear_regression.py
