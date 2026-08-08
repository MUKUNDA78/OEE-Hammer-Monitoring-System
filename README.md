# 🏭 GRS Engineering Pvt Ltd Unit 1 - OEE Monitoring & Analytics System

An industrial-grade, 5-hammer fleet efficiency monitoring and production analytics web application designed for forging plants. Features real-time visual OEE gauges, 12-hour shift system accounting, exact 17-column plant Excel imports with instant auto-commit, hammer-wise separate sheet uploads, part number analytics, and month-by-month trend performance tracking.

![OEE System Banner](https://img.shields.io/badge/System-OEE_Fleet_Monitor-16a34a?style=for-the-badge&logo=industry)
![License](https://img.shields.io/badge/License-MIT-blue?style=for-the-badge)
![Tech](https://img.shields.io/badge/Stack-HTML5_|_CSS3_|_JavaScript_|_Chart.js_|_SheetJS-0284c7?style=for-the-badge)

---

## 🌟 Key Features

### 🔨 1. Fleet Equipment Specifications (5 Hammers)
- **1 Ton Hammer** (Capacity: `1.0 Ton` | Standard Cycle: `35 sec`)
- **1.5 Ton Hammer** (Capacity: `1.5 Ton` | Standard Cycle: `42 sec`)
- **2.5 Ton (Old) Hammer** (Capacity: `2.5 Ton (Old)` | Standard Cycle: `55 sec`)
- **2.5 Ton (New) Hammer** (Capacity: `2.5 Ton (New)` | Standard Cycle: `48 sec`)
- **3.5 Ton Hammer** (Capacity: `3.5 Ton` | Standard Cycle: `70 sec`)

### ⏱️ 2. 12-Hour Shift System & Planned Operating Time Base
- **2 x 12-Hour Shifts**:
  - `Shift A`: `08:00 - 20:00`
  - `Shift B`: `20:00 - 08:00`
- **Break Accounting**: Excludes 30m Lunch + 30m Tea (2x15m) = 60m planned non-working breaks.
- **Net Planned Operating Base Time**: **660 Minutes (11.0 Hours)**.

### 📊 3. Industrial OEE Mathematical Engine
- **Availability ($A$)**:
  $$A = \frac{\text{Operating Time (660m - Total Downtime)}}{660\text{ Mins}} \times 100$$
- **Performance ($P$)**:
  $$P = \frac{(\text{Total Parts} \times \text{Ideal Cycle Sec}) / 60}{\text{Operating Time Mins}} \times 100$$
- **Quality ($Q$)**:
  $$Q = \frac{\text{Good Parts}}{\text{Total Parts}} \times 100$$
- **Overall OEE**:
  $$\text{OEE \%} = A \times P \times Q$$

### 📂 4. Exact 17-Column Plant Excel Import & Auto-Commit System
Supports exact plant Excel columns:
`Date`, `Shift`, `Machine`, `part number`, `Planned time`, `Maintance`, `die related`, `setup`, `No manpower`, `Heating time`, `minor stop`, `total downtime`, `operating time`, `total parts`, `good parts`, `rejects`, `ideal cycle time`.

- **Instant Auto-Commit**: Selecting or dropping an Excel file immediately parses, normalizes units, calculates OEE, saves logs, and switches to the Overview Dashboard.
- **Separate Hammer Excel Upload Cards**: Option for each hammer to upload its Excel sheet separately with dedicated cards, dropzones, log count badges, and template downloads.
- **Fuzzy Header Matcher & Unit Normalizer**:
  - Handles variations in column names (ignores spaces, symbols, parens, case).
  - Automatically converts hours (`11h`, `12h`, `1.5h`) to minutes.
  - Automatically converts cycle times (`< 3` min/pc $\rightarrow$ sec/pc, `> 300` pcs/hr $\rightarrow$ sec/pc).

### 🔍 5. Part Number & Machine Analysis
- Alphanumeric plant part numbers (`A1#21`, `W1#164`, `A4#07`, `C2#14`, `B3#88`, `M5#102`).
- Interactive Hammer Sub-Nav selector (`All 5 Hammers Combined`, `1 Ton`, `1.5 Ton`, `2.5T Old`, `2.5T New`, `3.5 Ton`).
- Tracks top setup times, die downtime, and furnace heating delays for each individual machine.

### 📅 6. Month-Wise Performance & Dynamic Month Selector
- **Dynamic Month Selector**: Top navigation filter allows selecting specific production months (`Jan 2026`, `Feb 2026`, `Jul 2026`, `All Months Combined`).
- **Month-Wise OEE Trend Chart**: Plots 6 trend lines (Overall Fleet Average + all 5 individual hammers).
- **Monthly Performance Matrix Table**: Month-by-month breakdown of shift count, total produced parts, overall OEE, and hammer-wise OEE.

---

## 🚀 How to Run Locally

1. Double click `start_app.bat` or run:
   ```powershell
   powershell -ExecutionPolicy Bypass -File server.ps1
   ```
2. Open your web browser at:
   ```
   http://localhost:8080
   ```
3. Or open `index.html` directly in any browser.

---

## 📁 Repository Structure

```
oee-hammer-monitor/
├── index.html       # Application HTML5 Layout & Views
├── styles.css       # Light Green / Mint Industrial Styling & Dark Mode
├── app.js           # Robust Excel Parser & Mathematical OEE Engine
├── server.ps1       # Local PowerShell Web Server
├── start_app.bat    # Windows 1-Click Server Launcher
└── README.md        # System Documentation
```

---

## 📜 License
Developed for Industrial Forge Plant Efficiency Monitoring. Released under MIT License.
