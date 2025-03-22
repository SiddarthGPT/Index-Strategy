from flask import Flask, render_template, request, send_file
import pandas as pd
import os
from datetime import datetime

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
RESULT_FILE = "output/Backtest_Result.xlsx"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs("output", exist_ok=True)

def categorize(cagr):
    if cagr < 0:
        return "Extreme Bearish"
    elif 0 <= cagr < 0.06:
        return "Bearish"
    elif 0.06 <= cagr < 0.10:
        return "Sideways Bearish"
    elif 0.10 <= cagr < 0.12:
        return "Neutral"
    elif 0.12 <= cagr < 0.15:
        return "Bullish"
    else:
        return "Extreme Bullish"

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files["file"]
        if file:
            filepath = os.path.join(UPLOAD_FOLDER, file.filename)
            file.save(filepath)

            df = pd.read_excel(filepath, sheet_name="Sheet1", skiprows=2, usecols="B:F", engine="openpyxl")
            df.columns = df.columns.str.strip()  # Clean column names
            df["Date"] = pd.to_datetime(df["Date"], errors="coerce")
            df = df.dropna(subset=["Date", "Close"]).reset_index(drop=True)

            cagr_records = []
            for i in range(len(df) - 250):
                entry_row = df.iloc[i]
                exit_row = df.iloc[i + 250]
                entry_price = entry_row["Close"]
                exit_price = exit_row["Close"]
                cagr = (exit_price / entry_price) ** (252 / 250) - 1
                cagr_records.append({
                    "Date": entry_row["Date"],
                    "Close": entry_price,
                    "Exit Date": exit_row["Date"],
                    "Exit Price": exit_price,
                    "Annualized CAGR": cagr,
                    "Category": categorize(cagr)
                })

            cagr_df = pd.DataFrame(cagr_records)

            capital = 2_500_000
            total_units = 0
            invested = 0
            withdrawn = 0
            cash = capital
            results = []

            for _, row in cagr_df.iterrows():
                date = row["Date"]
                price = row["Close"]
                cat = row["Category"]
                buy = sell = 0

                if cat == "Extreme Bearish" and cash >= 2 * price:
                    buy = 2
                elif cat == "Bearish" and cash >= price:
                    buy = 1
                elif cat == "Sideways Bearish" and cash >= 0.5 * price:
                    buy = 0.5

                invested += buy * price
                cash -= buy * price

                if cat == "Extreme Bullish" and total_units > 0.5:
                    sell = 1
                elif cat == "Bullish" and total_units > 0.5:
                    sell = 0.5

                withdrawn += sell * price
                cash += sell * price

                total_units += buy - sell
                portfolio_value = total_units * price

                results.append([
                    date, cat, price, buy, sell, total_units,
                    portfolio_value, invested, withdrawn, cash
                ])

            result_df = pd.DataFrame(results, columns=[
                "Date", "Category", "Close Price", "Units Bought", "Units Sold", "Total Units Held",
                "Portfolio Value", "Total Invested", "Total Withdrawn", "Remaining Cash"
            ])

            final_val = result_df.iloc[-1]["Portfolio Value"] + result_df.iloc[-1]["Remaining Cash"] + result_df.iloc[-1]["Total Withdrawn"]
            years = (result_df["Date"].max() - result_df["Date"].min()).days / 365.25
            cagr = (final_val / capital) ** (1 / years) - 1 if capital > 0 else 0

            summary = pd.DataFrame({
                "Metric": [
                    "Final Portfolio Value", "Remaining Cash", "Total Invested",
                    "Total Withdrawn", "Net Profit", "CAGR (on 25L)"
                ],
                "Value": [
                    result_df.iloc[-1]["Portfolio Value"],
                    result_df.iloc[-1]["Remaining Cash"],
                    result_df.iloc[-1]["Total Invested"],
                    result_df.iloc[-1]["Total Withdrawn"],
                    final_val - capital,
                    cagr
                ]
            })

            with pd.ExcelWriter(RESULT_FILE, engine="openpyxl") as writer:
                cagr_df.to_excel(writer, sheet_name="Historical Data with CAGR", index=False)
                result_df.to_excel(writer, sheet_name="Backtesting Results", index=False)
                summary.to_excel(writer, sheet_name="Performance Summary", index=False)

            return render_template("result.html")

    return render_template("index.html")

@app.route("/download")
def download():
    return send_file(RESULT_FILE, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
