# Currency Tracker

Tracks USD/EUR/CNY exchange rates from CBR (Bank of Russia) 
and saves history to Excel with change alerts.

## Features
- Live rates from CBR XML API
- Accumulates history in Excel (each run adds a row)
- Alerts when rate changes more than 0.1%

## Stack
- Python 3.x
- requests, openpyxl

## Run
pip install requests openpyxl
python currency_tracker.py
