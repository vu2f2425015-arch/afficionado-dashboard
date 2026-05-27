# ☕ Afficionado Coffee Roasters — Forecasting Dashboard

A coffee-themed Streamlit dashboard for **Data-Driven Forecasting & Peak Demand Prediction**.

## Features
- **Forecast tab** — daily sales forecast with ensemble of Naive, Moving Average, Exponential Smoothing, Linear Trend & Seasonal-Naive models, with 95% confidence band and backtest metrics (MAE / RMSE / MAPE).
- **Peak Demand** — day-of-week × hour heatmap, busiest hour & day callouts.
- **Stores** — per-store daily trend, revenue share donut, store × hour heatmap.
- **Products** — category & top-product revenue, category daily area trend.
- **Data** — filtered table + CSV export.
- **Quick-select filters** (`st.pills`) for store location & product category, date range slider, forecast horizon (1–30 days), Revenue/Quantity toggle.

## Run locally
```bash
pip install -r requirements.txt
streamlit run app.py
```

## Data
Reads `data/Afficionado_Coffee_Roasters.xlsx` directly (Transactions sheet).
The dataset has no explicit `transaction_date` column, so the app distributes
transactions in `transaction_id` order across Jan–Jun 2025 to enable daily
time-series forecasting. Hourly analysis uses the real `transaction_time`.
