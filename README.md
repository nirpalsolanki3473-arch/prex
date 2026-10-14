# 📊 Excel Formula Practice Workbook: Fundamental Booster

A beginner-friendly Excel project designed to strengthen fundamental spreadsheet skills through practical exercises using Student, Sales, and Employee datasets.

## 📌 Project Overview

This project is an Excel workbook designed to practice and understand commonly used Excel formulas through practical examples.

The workbook contains three worksheets — **Student, Sales, and Employee** — each representing a different type of dataset.

The project focuses on building a strong foundation in Excel formulas and applying them to structured data.

## 🎯 Objectives

- Practice essential Excel formulas and functions.
- Perform calculations on structured datasets.
- Apply conditional logic using `IF`, `AND`, and `OR`.
- Calculate totals, averages, and counts.
- Practice lookup functions to retrieve information.
- Manipulate text using Excel functions.
- Work with dates and numerical values.
- Understand formula references and cross-sheet lookups.

## 🛠️ Tools Used

- Microsoft Excel
- Excel formulas and functions
- Structured tabular data

The workbook can also be opened with compatible spreadsheet software, such as LibreOffice Calc, although some newer Excel functions may not be supported in every version.

## 📂 Workbook Structure

The workbook contains three worksheets.

### 1. Student

The Student worksheet contains student records and marks-related calculations.

**Columns include:**
- Student ID
- Name
- Math
- Science
- English
- Enrollment Date
- Total Marks
- Grade
- Math & Science
- Count Math (>50)
- Conditional average calculation
- Count based on total marks

**Excel concepts practiced:**
- `SUM` — calculate total marks
- `IF` — assign grades based on conditions
- `AND` — test multiple conditions
- `COUNTIF` — count values matching a condition
- `AVERAGE` — calculate average marks
- `COUNTIFS` — count records meeting criteria

### 2. Sales

The Sales worksheet contains sales transactions with product, region, salesperson, amount, and date information.

**Columns include:**
- Sales ID
- Product
- Region
- Salesperson
- Amount
- Date
- Capital
- Small
- Regional sales calculations
- Discount calculation
- Product lookup practice
- Eligibility check

**Excel concepts practiced:**
- `UPPER` — convert text to uppercase
- `LOWER` — convert text to lowercase
- `SUMIFS` — calculate totals using criteria
- `VLOOKUP` — retrieve matching information
- `IF` — calculate discounts based on price thresholds
- `OR` — check whether at least one condition is true
- `XMATCH` — find the position of a matching item

### 3. Employee

The Employee worksheet contains employee records and demonstrates salary lookups and text functions.

**Columns include:**
- Employee ID
- Name
- Department
- Salary
- Joining Date
- Lookup calculation
- Right
- Left

**Excel concepts practiced:**
- `XLOOKUP` — retrieve matching information
- `VLOOKUP` — retrieve data from another worksheet
- `LEFT` — extract characters from the beginning of text
- `RIGHT` — extract characters from the end of text
- `FLOOR` — round a number down to a specified multiple
- `CEILING` — round a number up to a specified multiple
- Cross-sheet references — retrieve information from another worksheet

## 📚 Key Excel Functions

| Function | Purpose |
|---|---|
| `SUM` | Adds numbers |
| `IF` | Returns results based on a condition |
| `AND` | Checks whether all conditions are true |
| `OR` | Checks whether at least one condition is true |
| `COUNTIF` | Counts cells matching one condition |
| `COUNTIFS` | Counts cells matching multiple conditions |
| `AVERAGE` | Calculates the average |
| `SUMIFS` | Adds values matching specified criteria |
| `VLOOKUP` | Looks up information vertically |
| `XLOOKUP` | Searches for a match and returns a corresponding value |
| `XMATCH` | Returns the position of a matching item |
| `UPPER` | Converts text to uppercase |
| `LOWER` | Converts text to lowercase |
| `LEFT` | Extracts characters from the left |
| `RIGHT` | Extracts characters from the right |
| `FLOOR` | Rounds down to a specified multiple |
| `CEILING` | Rounds up to a specified multiple |

## 💡 Example Formulas

### Calculate Total Marks

```excel
=SUM(C2:E2)
```

Adds the marks from Math, Science, and English.

### Assign Grades

```excel
=IF(G2>=220,"A",IF(G2>=190,"B",IF(G2>=160,"C","D")))
```

Assigns a grade based on the total marks.

### Count Marks Above 50

```excel
=COUNTIF(C2:C21,">50")
```

Counts the Math scores greater than 50.

### Calculate Discount

```excel
=IF(E2<10000,"0%",IF(E2<20000,"20%",IF(E2<30000,"30%","50%")))
```

Returns a discount category based on the sales amount.

### Extract Text

```excel
=LEFT(B2,8)
```

Extracts the first eight characters from the employee name.

```excel
=RIGHT(B2,2)
```

Extracts the last two characters from the employee name.

## ▶️ How to Use

1. Download `prex1.xlsx` from the repository.
2. Open it in Microsoft Excel or compatible spreadsheet software.
3. Navigate through the Student, Sales, and Employee worksheets.
4. Select cells containing formulas to inspect the formula bar.
5. Change sample values and observe how supported formulas respond.
6. Practice modifying formulas and applying them to additional rows.

**Note:** Formula results depend on the data, cell references, and Excel version. Review formulas and verify their outputs when using the workbook for practice.

## 🧠 Skills Demonstrated

- Excel formulas and functions
- Conditional calculations
- Logical functions
- Lookup and reference functions
- Text manipulation
- Date handling
- Basic data analysis
- Cross-sheet references
- Formula-based calculations

## 🚀 Future Improvements

- Add more sample datasets.
- Create summary tables and charts.
- Add pivot tables for sales and student analysis.
- Apply conditional formatting.
- Add data validation and dropdown lists.
- Create an interactive Excel dashboard.
- Add more advanced lookup and dynamic array formulas.

These are potential improvements and are not claimed as existing features of the workbook.

## 👨‍💻 Author

**NIRPALSINH SOLANKI**

GitHub: [@nirpalsolanki3473-arch](https://github.com/nirpalsolanki3473-arch)

## ⭐ Project Summary

This Excel workbook provides hands-on practice with common spreadsheet functions using student, sales, and employee data. It is suitable for beginners learning Excel formulas and developing practical spreadsheet skills for data analysis.
