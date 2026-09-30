# NextHikes_Project_1
# 🚲 Bike Sharing Demand Analysis Project Using Excel

Analysis of hourly bike-sharing data to understand how **weather, time of day, day type and holidays** affect rentals. Built with Excel (formulas, Pivot Tables, Slicers, Dashboard) and automated with **VBA**.

**Author:** MD SHAHNSHAH
**Organisation:** NextHikes IT Solutions
**Tools:** Microsoft Excel, Pivot Tables, Slicers, VBA

---

## 📌 Project Objective
Find patterns in bike rentals and turn them into business decisions:
- Who rides? (Casual vs Registered)
- When do they ride? (Weekday/Weekend, time of day)
- What affects demand? (Weather, temperature, humidity, wind)

## 📂 Repository Structure
```
├── NEXT HIKE IT SOLUTION PROJECT 1
  ├── Bike_Sharing_Demand_Analysis.pptx
  ├── NextHikes IT SOLUTION.xlsx
├── LICENSE          
└── README.md
```

## 📊 Dataset
| Item | Detail |
|---|---|
| Records | 1,000 hourly rows |
| Period | 1 Jan 2011 – 14 Feb 2011 (45 days) |
| Sources | 3 datasets merged on `instant` |
| Key fields | Date, Hour, Season, Weather, Temp, Humidity, Wind Speed, Casual, Registered, Count |

> This is an early 45-day sample, so results are an early signal, not a full-year conclusion.

## 🗂️ Workbook Sheets
| Sheet | Purpose |
|---|---|
| `DataSet_First / Second / Third` | Raw source data |
| `Merge_Data_Set` | Datasets merged on common key |
| `Final_Data_Set` | Cleaned table `FinaLData` with new columns (DayType, Time_Slot, Season_Name, Weather_Name, Holiday_Status, Temp_Types) |
| `Reference_Tables` | Lookup tables (Season, Weather, Temperature) |
| `Analysis` | COUNTIF/SUMIF/CORREL/STDEV statistics |
| `Pivot_Analysis` | Pivot tables and charts |
| `DASHBOARD` | Interactive dashboard with 6 slicers |
| `VBA` | KPI panel updated by macros |

## 🧹 Data Preparation
1. **Merge** – 3 datasets combined using `VLOOKUP / INDEX-MATCH`
2. **Clean** – duplicates removed, data types fixed, 137 missing WindSpeed values filled using `AVERAGEIFS`
3. **Feature engineering** – `IF` and `VLOOKUP` used to create DayType, Time_Slot, Season_Name, Weather_Name, Temp_Types

## 🤖 VBA Automation
| Macro | What it does |
|---|---|
| `Update_KPI` | Calculates Weekend / Weekday rentals and writes a "Last Updated" time |
| `Clean_Data` | Removes duplicate rows and counts blank WindSpeed cells (loop + condition) |
| `Check_KPI` | Verifies Casual + Registered = Total and Weekend + Weekday = Total |
| `Run_Full_Report` | One click: clean → refresh pivots → update KPIs → verify → open Dashboard |
| `Demand_Level()` | Custom function: High (100+), Medium (50+), Low |

**KPIs shown on the VBA sheet**

| KPI | Value |
|---|---|
| Total Records | 1,000 |
| Total Rentals | 58,304 |
| Average Rentals / hour | 58.304 |
| Casual Users | 4,921 |
| Registered Users | 53,383 |
| Weekend Rentals | 15,869 |
| Weekday Rentals | 42,435 |

### ▶️ How to run the macros
1. Download `NextHikes_IT_SOLUTION_VBA.xlsm`
2. Open it and click **Enable Content** in the yellow bar
3. Go to the `VBA` sheet and click the **Update Report** button
   (or `Developer → Macros → Run_Full_Report → Run`)

## 🔍 Key Findings
- **Registered riders make up 91.6%** of all rides – the business runs on commuters.
- **Clear weather = 64%** of rides; heavy rain/snow almost stops demand.
- **Weekdays** average 63.3 rides/hour vs ~48 on weekends.
- **Peak hours:** 8 AM and 5–6 PM (commute times).
- **Hot** conditions give ~68% more rides/hour than **Cold**.
- All rider segments trend upward over the 45-day period.

## 💡 Recommendations
1. Keep enough bikes available at 8 AM and 5–6 PM.
2. Convert casual riders to subscribers with targeted offers.
3. Plan bike placement around weather forecasts.
4. Re-check the trend with full-year data.


## ⚠️ Notes
- Macros only work in the `.xlsm` file, not `.xlsx`.
- Category names such as `WeakDay` are used exactly as spelled in the workbook.

## 📬 Contact
**MD SHAHNSHAH** –(https://www.linkedin.com/in/mdshahnshah).
