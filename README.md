# Excel Data Analyser & Transformer

### Interactive Excel-Based Data Analysis, Filtering & Transformation Tool

> "Data becomes useful when it can be organised, analysed, filtered, and transformed effectively."

---

## 📋 Table of Contents

* [📌 Overview](#-overview)
* [🎯 Problem Statement](#-problem-statement)
* [✨ Key Features](#-key-features)
* [🏗️ Project Structure](#-project-structure)
* [🔄 Project Workflow](#-project-workflow)
* [📥 Part A — Data Input & Formatting](#-part-a--data-input--formatting)
* [📊 Part B — Formulas & Functions](#-part-b--formulas--functions)
* [🛠️ Tech Stack](#-tech-stack)
* [📈 Results & Insights](#-results--insights)
* [🏆 Advantages](#-advantages)
* [📄 License](#-license)
* [👤 Author](#-author)
* [🙏 Acknowledgements](#-acknowledgements)

---

## 📌 Overview

The **Excel Data Analyser & Transformer** is a beginner-friendly Excel-based data analysis project that demonstrates important spreadsheet concepts such as **cell references**, **formatting**, **data input**, **IF formulas**, **logical conditions**, **lookup functions**, **date and time functions**, and **data filtering**.

The project is designed to:

* Practice data entry and spreadsheet formatting
* Understand relative and absolute cell references
* Apply IF and nested IF formulas for decision-making
* Use AND and OR conditions with IF formulas
* Apply COUNTIFS, SUMIFS, and AVERAGEIFS for conditional analysis
* Use TEXT functions for handling and displaying data
* Perform data lookup using VLOOKUP, INDEX & MATCH, XLOOKUP, and XMATCH
* Use INDIRECT and OFFSET functions for dynamic references
* Apply date, time, and mathematical functions
* Filter data using the FILTER function

---

## 🎯 Problem Statement

> **Objective:** Build an interactive Excel workbook that can accept, format, analyse, filter, and transform structured datasets using Excel formulas and functions.

The project contains multiple datasets such as **Students Grade**, **Sales Data**, and **Employee Data**. Excel formulas and functions are applied to these datasets to classify records, calculate results, perform lookups, apply conditions, and extract required information.

| 📂 Feature              | 📄 Type              | 🔍 Description                                   |
| ----------------------- | -------------------- | ------------------------------------------------ |
| Data Input              | Spreadsheet Input    | Stores student, sales, and employee information  |
| Cell References         | Formula Operation    | Uses relative and absolute references            |
| IF Analysis             | Conditional Analysis | Classifies records according to given conditions |
| Conditional Functions   | Data Analysis        | Uses COUNTIFS, SUMIFS, and AVERAGEIFS            |
| Lookup Operations       | Data Retrieval       | Retrieves information using lookup functions     |
| Filtering               | Data Filtering       | Extracts required records using FILTER           |
| Date & Time             | Date Analysis        | Works with date-related information              |
| Mathematical Operations | Calculation          | Performs required numerical calculations         |

The goal is to demonstrate **fundamental Excel data analysis skills** through structured datasets and formula-based operations.

---

## ✨ Key Features

| Feature                               | Description                                             |
| ------------------------------------- | ------------------------------------------------------- |
| 📊 **Data Input & Formatting**        | Allows structured data entry and spreadsheet formatting |
| 🔗 **Relative & Absolute References** | Uses cell references such as A1 and $A$1                |
| 🧠 **IF & Nested IF**                 | Applies multiple conditions for classification          |
| 🔀 **AND / OR Conditions**            | Combines multiple logical conditions                    |
| 🔢 **Conditional Functions**          | Uses COUNTIFS, SUMIFS, and AVERAGEIFS                   |
| 🔤 **TEXT Functions**                 | Processes and formats text-based information            |
| 🔎 **Lookup Functions**               | Uses VLOOKUP, INDEX & MATCH, XLOOKUP, and XMATCH        |
| 🔗 **Dynamic References**             | Uses INDIRECT and OFFSET for reference-based operations |
| 📅 **Date & Time Functions**          | Performs date and time-related calculations             |
| 🧮 **Math Functions**                 | Performs mathematical calculations on datasets          |
| 🔍 **FILTER Function**                | Returns multiple matching values from data              |

---

## 🏗️ Project Structure

```text
📦 excel-data-analyser-transformer/
│
├── 📊 Excel_Project.xlsx       ← Main Excel workbook
│
├── 📄 README.md                ← Project documentation
│
└── 📑 Sheets
    ├── Project Instructions
    ├── Students Grade
    ├── Sales Data
    └── Employee Data
```

---

## 🔄 Project Workflow

```text
Project Start
      │
      ▼
┌─────────────────────────────────┐
│       Open Excel Workbook       │
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│       Enter / Format Data       │
└──────────────┬──────────────────┘
               │
               ▼
┌─────────────────────────────────┐
│     Apply Excel Formulas        │
└──────────────┬──────────────────┘
               │
       ┌───────┼─────────────────────────────┐
       ▼       ▼             ▼               ▼
   ┌──────┐ ┌────────┐ ┌──────────┐ ┌────────────┐
   │  IF  │ │COUNTIFS│ │  Lookup  │ │   FILTER   │
   │Forms. │ │SUMIFS  │ │ Functions│ │  Function  │
   └──┬───┘ └───┬────┘ └────┬─────┘ └─────┬──────┘
      │         │            │             │
      └─────────┴────────────┴─────────────┘
                         │
                         ▼
                ┌──────────────────┐
                │ Analysed Results │
                └────────┬─────────┘
                         │
                         ▼
                ┌──────────────────┐
                │ Filter / Review  │
                │     Results      │
                └────────┬─────────┘
                         │
                         ▼
                    Final Output
```

---

## 📥 Part A — Data Input & Formatting

### 📝 1. What is Data Input?

The project contains structured datasets that are entered and organised inside Excel worksheets. Data includes student information, sales records, employee information, dates, scores, amounts, and other related values.

---

### 🗂️ 2. Workbook Sheets — Overview

| Sheet | Purpose                  |
| ----- | ------------------------ |
| 1️⃣   | **Project Instructions** |
| 2️⃣   | **Students Grade**       |
| 3️⃣   | **Sales Data**           |
| 4️⃣   | **Employee Data**        |

---

### 🔢 3. Relative & Absolute Cell References

> Relative references change when a formula is copied, while absolute references remain fixed.

Examples:

```text
Relative Reference: A1
Absolute Reference: $A$1
```

Absolute references are useful when a formula needs to repeatedly refer to a fixed value or parameter.

---

### 🎨 4. Formatting & Data Input

The workbook uses spreadsheet formatting to organise information clearly.

Examples include:

* Bold and formatted headings
* Date formatting
* Number and currency formatting
* Structured data entry
* Relative and absolute references

---

## 📊 Part B — Formulas & Functions

### 🧠 5. IF Formula & Nested IF

> Uses IF formulas to evaluate conditions and return appropriate results.

Nested IF formulas are used when multiple conditions need to be checked sequentially.

Example:

```excel
=IF(A1>=90,"A+",IF(A1>=80,"A",IF(A1>=70,"B","C")))
```

This can be used to classify student performance or apply different conditions to sales data.

---

### 🔀 6. IF with AND / OR

> Combines IF with logical conditions to analyse multiple requirements.

Example:

```excel
=IF(AND(C2>80,D2>80),"Yes","No")
```

AND requires all specified conditions to be satisfied.

OR returns TRUE when at least one of the specified conditions is satisfied.

---

### 🔢 7. COUNTIFS, SUMIFS & AVERAGEIFS

> These functions perform calculations based on one or more conditions.

| Function       | Purpose                                                         |
| -------------- | --------------------------------------------------------------- |
| **COUNTIFS**   | Counts records satisfying multiple conditions                   |
| **SUMIFS**     | Adds values satisfying multiple conditions                      |
| **AVERAGEIFS** | Calculates the average of values satisfying multiple conditions |

These functions can be used for analysing student, sales, or employee datasets.

---

### 🔤 8. Excel TEXT Functions

> TEXT functions are used to work with and display text and formatted values.

They can be applied to names, dates, and other text-based information in the workbook.

---

### 🔎 9. VLOOKUP Function

> VLOOKUP searches for a value in the first column of a table and returns related information from another column.

It is useful for retrieving information from structured datasets.

---

### 🔗 10. INDEX & MATCH Functions

> INDEX returns a value from a specified position, while MATCH identifies the position of a value.

Together, they can be used for flexible data lookup operations.

---

### 🔍 11. XLOOKUP & XMATCH Functions

> XLOOKUP searches for a value and returns the corresponding result, while XMATCH identifies the position of a matching value.

These functions provide flexible lookup and matching operations within Excel data.

---

### 🔗 12. INDIRECT Function

> INDIRECT converts a text reference into an actual cell or range reference.

It can be used when the cell reference needs to be generated dynamically.

---

### ↔️ 13. OFFSET Function

> OFFSET returns a reference that is a specified number of rows and columns away from a starting cell.

It can be used for dynamic reference-based operations.

---

### 📅 14. Date & Time Functions

> Date and time functions are used to work with dates and time-related values in datasets.

They can help analyse employee joining dates, sales dates, and other date-based information.

---

### 🧮 15. Maths Functions

> Mathematical functions are used to perform calculations and manipulate numerical values.

They can be applied to student marks, sales amounts, salaries, and other numerical datasets.

---

### 🔍 16. FILTER Function

> FILTER returns multiple values that satisfy a specified condition.

It is useful for extracting only the required records from a larger dataset.

Example:

```excel
=FILTER(A2:E1000,C2:C1000>80)
```

---

## 🛠️ Tech Stack

| Tool                                  | Version                  | Purpose                           |
| ------------------------------------- | ------------------------ | --------------------------------- |
| 📊 **Microsoft Excel**                | Excel 365 / Modern Excel | Spreadsheet and data analysis     |
| 🔗 **Cell References**                | Built-in                 | Relative and absolute referencing |
| 🧠 **IF Functions**                   | Built-in                 | Conditional analysis              |
| 🔢 **COUNTIFS / SUMIFS / AVERAGEIFS** | Built-in                 | Conditional calculations          |
| 🔎 **Lookup Functions**               | Built-in                 | Data retrieval and matching       |
| 📅 **Date & Time Functions**          | Built-in                 | Date-related calculations         |
| 🧮 **Math Functions**                 | Built-in                 | Numerical calculations            |
| 🔍 **FILTER**                         | Built-in                 | Filtering and extracting data     |

---

## 📈 Results & Insights

After applying the required Excel formulas and functions, the workbook provides:

* ✅ **Structured Data** — Student, sales, and employee information organised into separate worksheets
* 📊 **Conditional Analysis** — Records classified using IF, AND, and OR conditions
* 🔢 **Conditional Calculations** — COUNTIFS, SUMIFS, and AVERAGEIFS applied for data analysis
* 🔎 **Data Lookup** — Required information retrieved using lookup functions
* 📅 **Date Analysis** — Date and time information processed using relevant functions
* 🔍 **Data Filtering** — Required records extracted using the FILTER function
* 🧮 **Mathematical Processing** — Numerical values processed using mathematical functions

---

## 🏆 Advantages

| Advantage                | Detail                                                         |
| ------------------------ | -------------------------------------------------------------- |
| 🎓 **Beginner Friendly** | Introduces important Excel concepts through practical datasets |
| 🔄 **Reusable**          | Formulas can be copied and applied to additional records       |
| 📚 **Educational**       | Each worksheet reinforces a different Excel concept            |
| 📊 **Structured**        | Data is organised into separate worksheets                     |
| ⚡ **Efficient**          | Excel formulas reduce repetitive manual calculations           |
| 🔍 **Flexible**          | Lookup and filtering functions allow targeted data retrieval   |
| 📈 **Practical**         | Uses student, sales, and employee datasets for analysis        |
| 🧮 **Data Analysis**     | Helps transform raw spreadsheet data into useful results       |

---

## 📄 License

This project is licensed under the **MIT License**.

```text
MIT License — Free to use, modify, and distribute with attribution.
```

---

## 👤 Author

**Gosai Kishal**

> "Good data becomes meaningful when it is organised, analysed, and understood."

**🎓 Role:** B.Tech Student | Excel & Data Analysis Learner

**📍 Location:** India

**🛠️ Skills:** Excel · Data Analysis · Formulas · Functions · Data Filtering · Lookup Functions

---

## 🙏 Acknowledgements

Special thanks to the following resources and communities that made this project possible:

* 📚 Microsoft Excel — Spreadsheet and data analysis platform
* 📖 Excel Documentation — Function and formula references
* 🎓 Course Materials — Excel concepts and practical exercises
* 💻 Excel Learning Resources — Formula and data analysis examples

---

*Made with ❤️ and ☕ — Last updated: 01 October, 2026*
