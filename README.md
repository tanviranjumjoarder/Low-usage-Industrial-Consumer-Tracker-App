# ⚡ Low Consumption Consumer Tracker

> A desktop utility for electricity distribution offices to identify industrial consumers with abnormally low unit consumption — helping engineers and billing officers detect billing anomalies, meter faults, or illegal connections, and take action to reduce system loss.

![Python](https://img.shields.io/badge/Python-3.8%2B-blue?logo=python)
![Platform](https://img.shields.io/badge/Platform-Windows-lightgrey?logo=windows)
![License](https://img.shields.io/badge/License-MIT-green)
![Language](https://img.shields.io/badge/UI%20Language-Bengali%20%2F%20English-orange)

---

## 📋 Table of Contents

- [Background & Purpose](#-background--purpose)
- [Features](#-features)
- [Screenshots / How It Works](#-how-it-works)
- [Input File Format](#-input-file-format)
- [Office Codes](#-office-codes)
- [Installation & Running](#-installation--running)
- [Building a Standalone EXE](#-building-a-standalone-exe)
- [Output](#-output)
- [Project Structure](#-project-structure)
- [Contributing](#-contributing)
- [License](#-license)

---

## 🏭 Background & Purpose

In electricity distribution networks, **system loss** is a major operational challenge. One contributing factor is industrial consumers who report suspiciously low unit consumption month after month — potentially indicating:

- **Faulty or tampered meters**
- **Illegal bypass connections**
- **Billing data entry errors**
- **Inactive consumers still on active accounts**

This tool was built for use in the **Kushtia (Natore/Rajshahi Division)** distribution zone. Each month, billing staff export consumer data from their billing system into Excel files. This app:

1. **Merges** multiple monthly Excel files into one consolidated dataset
2. **Splits** consumers by their issuing office (10 offices supported)
3. **Filters** consumers whose usage fell at or below a configurable threshold (default: **150 kWh**) in any month
4. **Exports** one formatted Excel report per office for field follow-up

The goal is to give supervisors and engineers a quick, actionable list of consumers to physically inspect — helping reduce system loss and improve revenue recovery.

---

## ✨ Features

- 🖥️ **Clean GUI** — no command-line knowledge needed; built with Python's `tkinter`
- 📂 **Multi-file merge** — load any number of monthly `.xlsx` billing export files at once
- 🏢 **Office-wise split** — automatically separates consumers into 10 sub-office files based on their account number prefix
- 🔻 **Low-usage filter** — keep only consumers with ≤ N kWh in any tracked month (threshold is adjustable)
- 📊 **Formatted Excel output** — colour-coded headers, alternating row fills, frozen panes, proper column widths
- 🗓️ **Auto month detection** — reads the billing month from the data or falls back to the filename
- 📱 **Mobile number cleaning** — normalises phone numbers, fills missing as `N/A`
- 🌐 **Bilingual** — Bengali column headers and UI labels, English log output
- 🔁 **Missing months filled** — consumers absent in a month get `0` (not blank) so filtering works correctly
- 📦 **Portable** — can be packaged into a single `.exe` with no Python dependency on the target machine

---

## 🔍 How It Works

```
Monthly billing exports (Excel)
         │
         ▼
┌─────────────────────┐
│   Load & Validate   │  ← checks required columns, detects billing month
└─────────────────────┘
         │
         ▼
┌─────────────────────┐
│  Merge All Months   │  ← pivot: one row per consumer, one column per month
└─────────────────────┘
         │
         ▼
┌─────────────────────┐
│  Split by Office    │  ← matches first 6 digits of account number to office code
└─────────────────────┘
         │
         ▼
┌─────────────────────┐
│  Apply Low-Usage    │  ← optional: drop consumers who never hit the threshold
│  Filter (≤ N kWh)   │
└─────────────────────┘
         │
         ▼
┌─────────────────────┐
│  Write Styled XLSX  │  ← one file per office, in output subfolder
└─────────────────────┘
```

---

## 📄 Input File Format

Each monthly Excel file must have at least these Bengali column headers (exact spelling required):

| Column Header (Bengali) | Description |
|---|---|
| `হিসাব নম্বর` | Consumer account number (first 6 digits = office code) |
| `গ্রাহকের নাম` | Consumer name |
| `মোট ব্যবহৃত (কি.ও.ঘ.)` | Total units consumed (kWh) |

These columns are **optional but recommended**:

| Column Header | Description |
|---|---|
| `ট্যারিফ` | Tariff category |
| `চুক্তিবদ্ধ লোড (কি.ও.)` | Contracted load (kW) |
| `বিল মাস` | Billing month (date column — used to auto-detect month label) |
| `মোবাইল নম্বর` | Consumer mobile number |

> **Tip:** If your billing software exports column names slightly differently (e.g. `হিসাব_নম্বর` with underscore), the app has built-in alias matching for common variants. See `ALIASES` in `merge_app.py` to extend this.

The billing month can also be detected from the **filename** if it follows the pattern:  
`...January 2024...`, `...March_2025...`, `...february-24...` etc.

---

## 🏢 Office Codes

The app recognises 10 sub-offices by the **first 6 digits** of the consumer account number:

| Code | Office Name |
|------|-------------|
| 101901 | HQ |
| 101902 | Daulotpur |
| 101903 | Kumarkhali |
| 101904 | Mirpur |
| 101905 | Bheramara |
| 101906 | Swastipur |
| 101907 | Khoksha |
| 101908 | Poradah |
| 101909 | Pragpur |
| 101910 | Panti |

To adapt this tool for a different distribution zone, edit the `OFFICES` dictionary at the top of `merge_app.py`.

---

## 🚀 Installation & Running

### Prerequisites

- **Python 3.8 or later** — download from [python.org](https://www.python.org/downloads/)  
  ✅ During installation, tick **"Add Python to PATH"**

### Option 1 — Run directly (recommended for development)

```bash
# Clone the repository
git clone https://github.com/YOUR_USERNAME/low-consumption-consumer-tracker.git
cd low-consumption-consumer-tracker

# Install dependencies
pip install pandas openpyxl

# Launch the app
python merge_app.py
```

> The app will auto-install `pandas` and `openpyxl` on first launch if they are missing.

### Option 2 — Double-click launcher (Windows)

1. Put all project files in one folder.
2. Double-click **`Run_Merger.bat`** (or **`Low_Consumption_Consumer_Tracker.bat`**).
3. The GUI opens. No terminal window stays open if `pythonw` is available.

---

## 📦 Building a Standalone EXE

To create a single `.exe` that runs without Python installed:

1. Ensure Python 3.8+ is installed and on PATH.
2. Place `app_icon.ico` in the same folder as the other files (optional — provides the lightning bolt icon).
3. Double-click **`Build_EXE.bat`**.
4. Wait 1–2 minutes. The finished app appears at:
   ```
   dist\Electricity Bill Merger.exe
   ```
5. Copy that single file anywhere — USB stick, another PC, shared drive — and double-click to run.

The build script uses [PyInstaller](https://pyinstaller.org/) and installs it automatically if needed.

---

## 📁 Output

After processing, the app creates a subfolder called:

```
Low usage Industry Consumer Officewise/
├── HQ.xlsx
├── Daulotpur.xlsx
├── Kumarkhali.xlsx
├── Mirpur.xlsx
├── Bheramara.xlsx
├── Swastipur.xlsx
├── Khoksha.xlsx
├── Poradah.xlsx
├── Pragpur.xlsx
└── Panti.xlsx
```

Each file contains a formatted table with:
- Consumer account number, name, tariff, contracted load
- One column per billing month showing kWh consumed
- Mobile number (last column)
- Only consumers who triggered the low-usage filter (if filtering is enabled)

The output folder opens automatically in Windows Explorer when processing completes.

---

## 🗂️ Project Structure

```
low-consumption-consumer-tracker/
│
├── merge_app.py                    # Main application (processing engine + GUI)
├── Run_Merger.bat                  # Windows double-click launcher
├── Low_Consumption_Consumer_Tracker.bat  # Alternative launcher
├── Build_EXE.bat                   # PyInstaller build script
├── HOW_TO_RUN.txt                  # End-user quick-start instructions
├── app_icon.ico                    # Application icon (not included — add your own)
│
├── README.md                       # This file
├── CONTRIBUTING.md                 # How to contribute
├── CHANGELOG.md                    # Version history
└── LICENSE                         # MIT License
```

---

## 🤝 Contributing

Contributions are welcome! See [CONTRIBUTING.md](CONTRIBUTING.md) for guidelines.

Ideas for improvement:
- Support for additional office codes / zones
- Exportable PDF summary report
- Chart/graph of monthly usage trends per consumer
- Email or SMS alert integration for flagged consumers
- Automated scheduling (run on a set date each month)

---

## 📜 License

This project is licensed under the **MIT License** — see [LICENSE](LICENSE) for details.

---

## 👤 Author

Built by a distribution engineer to solve a real operational problem in electricity billing management.  
If this tool helps your office, a ⭐ on GitHub is appreciated!
