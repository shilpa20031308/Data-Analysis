# 📊 Data Analysis Project using Microsoft Excel

This project contains multiple **Data Analysis case studies** created using **Microsoft Excel**.  
Each dataset demonstrates how Excel can be used to organize, calculate, visualize, and draw insights from real-world data.

---

## 📁 Project Overview

This repository includes 5 different Excel-based data analysis projects:

1. **Payroll Analysis**
2. **Gradebook Analysis**
3. **Decision Factors Analysis**
4. **Sales Database Analysis**
5. **Car Inventory Analysis**

Each sheet focuses on:
- Data Cleaning and Formatting  
- Use of Formulas and Functions  
- Charts and Visualization  
- Summary Reports and Insights  

---

## 🧠 Tools Used
- **Microsoft Excel**
- **Functions:** `SUM()`, `AVERAGE()`, `IF()`, `VLOOKUP()`, `HLOOKUP()`, `COUNTIF()`, `ROUND()`, `SUMPRODUCT()`, etc.  
- **Excel Charts:** Bar, Pie, Line, Column, and Pivot Charts  
- **Excel Features:** Conditional Formatting, Data Validation, Pivot Tables, and Dashboards  

---

## ⚙️ Excel Functions Summary

| Function | Purpose | Example |
|-----------|----------|----------|
| `SUM()` | Adds values in a range | `=SUM(B2:B10)` |
| `AVERAGE()` | Calculates average of values | `=AVERAGE(C2:C10)` |
| `IF()` | Applies condition-based logic | `=IF(D2>60,"Pass","Fail")` |
| `VLOOKUP()` | Searches a value vertically in a table | `=VLOOKUP(A2,Sheet2!A:B,2,FALSE)` |
| `HLOOKUP()` | Searches a value horizontally | `=HLOOKUP("Math",A1:F5,2,FALSE)` |
| `COUNTIF()` | Counts cells matching criteria | `=COUNTIF(C2:C20,">=50")` |
| `ROUND()` | Rounds a number to specific decimals | `=ROUND(E2,2)` |
| `SUMPRODUCT()` | Multiplies and sums corresponding ranges | `=SUMPRODUCT(B2:B10,C2:C10)` |
| `MAX()` / `MIN()` | Finds highest / lowest value | `=MAX(E2:E20)` |
| `ABS()` | Returns absolute value | `=ABS(D2)` |
| `RANK()` | Ranks data in order | `=RANK(F2,F2:F20,0)` |
| `CONCAT()` | Joins text strings | `=CONCAT(A2," ",B2)` |

---

## 1️⃣ Payroll Data Analysis

**Goal:** To analyze employee salary details and generate insights such as gross pay, deductions, and net salary.

### 🔹 Key Formulas:
- **Gross Salary:** `=Basic_Pay + Allowances - Deductions`
- **Tax Calculation:** `=Gross_Salary * Tax_Rate`
- **Net Salary:** `=Gross_Salary - Tax`
- **Average Salary:** `=AVERAGE(Salary_Range)`
- **Highest Salary:** `=MAX(Salary_Range)`

### 🔹 Insights:
- Salary distribution by department  
- Total payroll expenses  
- Employees with highest and lowest pay  
<p>
📁 <strong>Project 1: Payroll Analysis</strong><br>
<a href="https://github.com/shilpa20031308/Data-Analysis/blob/main/Project_1.xlsx" target="_blank">
Click here to open Project_1.xlsx
</a>
</p>


## 2️⃣ Gradebook Data Analysis

**Goal:** Evaluate student marks and calculate overall grades and performance summaries.

### 🔹 Key Formulas:
- **Total Marks:** `=SUM(Mark1:MarkN)`
- **Average Marks:** `=AVERAGE(Mark1:MarkN)`
- **Percentage:** `=(Total_Marks / Maximum_Marks) * 100`
- **Grade:**  
  ```excel
  =IF(Percentage>=90,"A",
     IF(Percentage>=80,"B",
     IF(Percentage>=70,"C",
     IF(Percentage>=60,"D","F"))))

### 🔹Insights:
-  Subject-wise and student-wise performance

-  Pass/fail summary using COUNTIF()

-  Top performers ranking
<p>
📁 <strong>Project 2: Gradebook Analysis</strong><br>
<a href="https://github.com/shilpa20031308/Data-Analysis/blob/main/Project_2.xlsx" target="_blank">
Click here to open Project_2.xlsx
</a>
</p>
  

##  3️⃣ Decision Factors Analysis

**Goal:** Identify and prioritize business decision factors using weighted scoring.

### 🔹 Key Formulas:

-  **Weighted Score**: =Factor_Value * Weight

-  **Total Weighted Score**: =SUMPRODUCT(Factor_Range, Weight_Range)

-  **Decision Index**: =(Weighted_Score / Total_Weight) * 100

### 🔹 Insights:

-  Rank of key decision factors

-  Comparison of alternatives

-  Support for data-driven business planning
  <p>
📁 <strong>Project 3: Decision Factors</strong><br>
<a href="https://github.com/shilpa20031308/Data-Analysis/blob/main/Decision%20Marker.xlsx" target="_blank">
Click here to open Decision_Marker.xlsx
</a>
</p>

   ## 4️⃣ Sales Database Analysis

**Goal:** Analyze sales data to determine revenue, profit, and sales trends.

### 🔹 Key Formulas:

-  **Revenue**: =Units_Sold * Selling_Price

-  **Profit**: =Revenue - Cost

-  **Profit Margin (%)**: =(Profit / Revenue) * 100

-  **Total Sales**: =SUM(Revenue_Range)

#### 🔹 Insights:

-  Best-selling products and regions

-  Monthly/Yearly sales trends using charts

-  Salesperson performance comparison

### 🔹 Visuals:

-  Line chart for sales trends

-  Pie chart for product contribution

-  Bar chart for profit by region
  <p>
📁 <strong>Project 4: Sales Database</strong><br>
<a href="https://github.com/shilpa20031308/Data-Analysis/blob/main/Sale%20Report.xlsx" target="_blank">
Click here to open Sale_Report.xlsx
</a>
</p>

## 5️⃣ Car Inventory Analysis

**Goal**: Manage and analyze car stock, pricing, and demand using Excel functions.

### 🔹 Key Formulas:

-  **Depreciation**: =(Initial_Price - Current_Price) / Years_Used

-  **Average Price**: =AVERAGE(Price_Range)

-  **Inventory Turnover**: =Cars_Sold / Average_Inventory

-  **Stock Status**:

=IF(Stock>10,"In Stock","Low Stock")

### 🔹 Insights:

-  Brand-wise car pricing

-  Most demanded car models

-  Depreciation comparison across brands
  <p>
📁 <strong>Project 5: Car Inventory</strong><br>
<a href="https://github.com/shilpa20031308/Data-Analysis/blob/main/car%20inventory%20(1).xlsx" target="_blank">
Click here to open car inventory (1).xlsx
</a>
</p>

 ### 📈 Visualization & Reports

- **Each dataset includes**:

-  Pivot Tables for summary and insights

-  Charts (Bar, Line, and Pie) to visualize key results

-  Conditional Formatting to highlight trends (e.g., high profits, low grades)

- Dashboards summarizing important metrics

  ## 📊 Project Results

| Dataset          | Focus Area            | Key Outcome                                       |
| ---------------- | --------------------- | ------------------------------------------------- |
| Payroll          | Employee Pay Analysis | Identified salary trends and payroll distribution |
| Gradebook        | Student Performance   | Ranked students and generated grading reports     |
| Decision Factors | Business Criteria     | Weighted decision matrix for strategy             |
| Sales Database   | Revenue & Profit      | Monthly trends and top-performing products        |
| Car Inventory    | Stock Management      | Optimized car pricing and turnover rate           |


## 🧾 Conclusion

-  **This project showcases how Microsoft Excel can be a powerful tool for**:

- Business and academic data analysis

- Decision-making based on numerical and visual insights

- Generating professional reports without programming

 ## 👩‍💻 Author

## Ambothu Shilpa
- **📧 Email**: [shilpaambothu@gmail.com]
- **💼 GitHub**: [https://github.com/shilpa20031308]
-**🌐 LinkedIn**: [https://www.linkedin.com/in/ambothu-shilpa-1a907b259]
