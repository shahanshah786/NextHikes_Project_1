
# 🚲 Bike Sharing Demand Analysis Using Excel

An Excel-based data analytics project to explore bike-sharing demand and understand how temperature, weather conditions, time of day, and weekdays or weekends relate to rental patterns.

**Author:** MD SHAHNSHAH  
**Organisation:** NextHikes IT Solutions  
**Tools:** Microsoft Excel, Excel Formulas, Lookup Functions, Pivot Tables, Charts, Statistical Analysis, Regression Analysis

---

## 📌 Project Objective

The objective of this project is to analyze bike-sharing rental data and identify useful patterns that can support business decisions.

Key questions explored:

- How many bikes are rented in total?
- How do casual and registered users differ?
- How does rental demand vary by hour and day type?
- How are weather conditions and temperature related to bike rentals?
- Can statistical analysis help explain rental demand?

## 📂 Project Structure

```text
NextHikes_Project_1/
│
├── NEXT HIKE IT SOLUTION PROJECT 1/
│   ├── Bike_Sharing_Demand_Analysis.pptx
│   └── NextHikes IT SOLUTION.xlsx
│
├── LICENSE
└── README.md
```

## 📊 Dataset Overview

The workbook contains **1,000 hourly records** covering **1 January 2011 to 14 February 2011**.

Three source datasets were combined using the common `instant` identifier.

| Item | Details |
|---|---|
| Records | 1,000 hourly observations |
| Date range | 1 Jan 2011 – 14 Feb 2011 |
| Data sources | Three source datasets |
| Target variable | `Count` |
| User categories | `Casual`, `Registered` |
| Weather variables | `Temp`, `Atemp`, `Humidity`, `WindSpeed` |
| Time variables | `DateDay`, `Hour`, `WeekDay`, `Holiday` |

*Note: This is a limited sample, so the findings should not be treated as full-year or universal trends.*

## 🗂️ Workbook Sheets

| Sheet | Purpose |
|---|---|
| `DataSet_First` | First source dataset |
| `DataSet_Second` | Additional weather and rental information |
| `DataSet_Third` | Third source dataset |
| `Merge_Data_Set` | Combined source data |
| `Final_Data_Set` | Main dataset containing 1,000 records and derived columns |
| `Reference_Tables` | Lookup values for seasons, weather, and temperature categories |
| `Analysis` | Summary statistics, correlations, and analysis formulas |
| `ANOVA_SLR` | Simple Linear Regression and ANOVA output |
| `Pivot_Analysis` | Pivot-based summaries of rental demand |
| `Forcasting` | Daily rental values and forecast-related columns |
| `DASHBOARD` | Dashboard sheet reserved for dashboard development |
| `VBA` | KPI summary panel using Excel formulas |

## 🧹 Data Preparation

The following data preparation work was carried out in Excel:

1. **Data merging:** Combined source datasets using the common `instant` identifier.
2. **Lookup operations:** Used lookup-based techniques to bring related columns together.
3. **Data cleaning:** Worked with missing WindSpeed values and prepared a filled WindSpeed column.
4. **Feature engineering:** Added derived fields such as `DayType`, `Time_Slot`, `Season_Name`, `Weather_Name`, `Holiday_Status`, and `Temp_Types`.
5. **Reference tables:** Used category mappings to make the dataset easier to analyze.

## 📈 Data Analysis

The `Analysis` sheet includes calculations and statistical techniques such as:

- `COUNT` and `AVERAGE`
- `SUM` and `SUMIF`
- `COUNTIF`
- `STDEV`
- `CORREL`
- Comparisons involving temperature, humidity, wind speed, and rental count

These calculations help summarize the data and explore relationships between variables.

## 📊 Pivot Table Analysis

The `Pivot_Analysis` sheet contains summaries of rental activity by:

- Hour and day type
- Weather category
- Time slot
- Casual and registered users

These summaries help compare demand across different periods and categories.

## 📉 Simple Linear Regression (SLR) and ANOVA

The `ANOVA_SLR` sheet contains a regression analysis examining the relationship between the temperature-derived predictor and bike rental count.

Reported results in the workbook:

| Metric | Result |
|---|---:|
| Observations | 1,000 |
| R Square | 0.0526 |
| F-statistic | 55.38 |
| Significance F | 2.14 × 10⁻¹³ |
| Predictor coefficient | 3.1494 |

The reported model has a statistically significant relationship, while its R Square is approximately **5.26%**. This means the model explains only a small portion of the variation in rental counts, so additional predictors would be needed for a stronger model.

## 🔢 KPI Summary

The workbook's `VBA` sheet contains formulas for the following KPIs:

| KPI | Value |
|---|---:|
| Total Records | 1,000 |
| Total Rentals | 58,304 |
| Average Rentals per Hour | 58.304 |
| Casual Users | 4,921 |
| Registered Users | 53,383 |
| Weekend Rentals | 15,869 |
| Weekday Rentals | 42,435 |

These values summarize rental activity across the records in the workbook.

## 🔍 Key Findings

- Registered users account for approximately **91.6%** of recorded rentals.
- Weekday rentals total **42,435**, compared with **15,869** on weekends.
- The average rental count across all 1,000 hourly records is **58.304**.
- Pivot summaries allow rental activity to be compared by time slot and weather category.
- The current single-predictor regression has a low R Square, indicating that temperature alone does not explain most rental-count variation.

## 💡 Business Recommendations

1. Review hourly rental patterns to plan bike availability during busy periods.
2. Explore promotions that may encourage casual users to become registered users.
3. Evaluate weather-related patterns when planning bike distribution.
4. Add more predictors to future regression models.
5. Validate findings using a larger, preferably full-year dataset.

## 🛠️ VBA and Automation Status

The workbook currently includes a `VBA` worksheet with KPI formulas.

**Current file limitation:** The provided workbook is an `.xlsx` file. It does not store VBA macros. To implement macro-based automation, save a separate copy as `.xlsm` and add and test the required VBA procedures.

Potential automation tasks include:

- Updating KPI summaries
- Refreshing Pivot Tables
- Checking consistency between rental totals and user totals
- Checking missing values
- Running a report update from a button

These should be described as implemented features only after the corresponding macros have been added and tested.

## ⚠️ Limitations

- The dataset covers only a short period of approximately 45 days.
- The regression model has limited explanatory power.
- Results may not represent seasonal or full-year rental patterns.
- The `DASHBOARD` sheet still needs to be populated if an interactive dashboard is intended.
- VBA automation requires a macro-enabled `.xlsm` workbook.

## 📚 Skills Demonstrated

- Excel data cleaning and preparation
- Lookup functions and feature engineering
- Statistical summaries and correlation analysis
- Pivot Table analysis
- Simple Linear Regression and ANOVA
- KPI reporting
- Data interpretation and business recommendations

## 📬 Contact

**MD SHAHNSHAH**  
LinkedIn: https://www.linkedin.com/in/mdshahnshah
