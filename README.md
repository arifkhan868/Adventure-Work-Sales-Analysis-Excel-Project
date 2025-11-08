## 📊 AdventureWorks Sales Analysis Dashboard
### *An Advanced Excel BI Project by Ariful Islam*

![Excel](https://img.shields.io/badge/Excel-217346?style=for-the-badge&logo=microsoft-excel&logoColor=white)
![PowerQuery](https://img.shields.io/badge/Power_Query-blue?style=for-the-badge)
![PowerPivot](https://img.shields.io/badge/Power_Pivot-gray?style=for-the-badge)
![DAX](https://img.shields.io/badge/DAX-Data_Modeling-orange?style=for-the-badge)
![VBA](https://img.shields.io/badge/VBA-Automation-8B0000?style=for-the-badge)

---

### 🧠 **Project Overview**
The **AdventureWorks Sales Analysis Dashboard** is an **end-to-end Excel BI project** designed to deliver actionable business insights through advanced Excel features like Power Query, Power Pivot, DAX, and VBA.

It analyzes **global sales performance from 2005–2008**, focusing on revenue, profit margin, transactions, and key customer-product trends — all inside Excel, achieving near **Power BI-level interactivity**.

---

### 🎯 **Business Problem & Solution**

#### 🧩 Problem
AdventureWorks needed a way to **analyze global sales data** across different years, regions, and product categories to understand:
- What drives profitability?
- Which countries and customers perform best?
- How do price, color, and time affect sales trends?

#### 💡 Solution
Built a **fully automated Excel BI dashboard** integrating:
- Power Query for ETL (Extract, Transform, Load)
- Power Pivot for star-schema data modeling
- DAX for dynamic measures
- VBA Macros for one-click refresh and filter clearing
- Interactive slicers for user-driven analysis

---

### ⚙️ **Tools & Technologies**
| Category | Tools / Techniques Used |
|-----------|--------------------------|
| Data Cleaning | Power Query |
| Data Modeling | Power Pivot, Star Schema |
| Measures & KPIs | DAX (SUM, DIVIDE, CALCULATE, FILTER) |
| Visualization | Pivot Table, Slicers, Charts |
| Automation | VBA Macro Buttons |
| Business Logic | IF, IFS, VLOOKUP, Conditional Formatting |

---

### 📈 **Key KPIs & Metrics**

| KPI | Value | Description |
|------|------:|-------------|
| **Total Quantity** | 631.92K | Total units sold |
| **Total Revenue** | $307.09M | Overall sales revenue |
| **Total Profit** | $126.29M | After deducting COGS |
| **Profit Margin** | 41.1% | Profitability ratio |
| **Transactions** | 60.4K | Total unique orders |
| **Top Countries** | 🇺🇸 USA, 🇦🇺 Australia | 62.7% of total profit |
| **Top Products** | Mountain-200 Series | Most profitable items |

---

### 🗂️ **Dashboard Structure**

#### 🔷 **Time Analysis Dashboard**
- Yearly, Monthly, and Weekly Profit Trends
- Weekday vs Weekend contribution
- Quarterly performance overview
- KPI cards (Revenue, COGS, Profit, Margin, Transactions)
- Filter buttons for Year & Month

#### 🔶 **Detailed Analysis Dashboard**
- Top 5 Products & Customers by Profit
- Profit distribution by Gender, Age Group, and Region
- Product color and pricing insights
- Geo map with country-wise profit contribution
- Dynamic filtering by Country and Year

---

### 💻 **Dashboard Preview**

#### ⏱️ *Time Analysis View*
![Time Analysis Dashboard](Dashboard1.png)

#### 🧩 *Detail Dashboard View*
![Detail Dashboard](Dashboard2.png)

---

### 🔍 **Core DAX Measures Used**

```DAX
Total Revenue = SUM(Sales[SalesAmount])
Total COGS = SUM(Sales[Cost])
Total Profit = [Total Revenue] - [Total COGS]
Profit Margin % = DIVIDE([Total Profit], [Total Revenue])
```

---

### ⚙️ **Automation with VBA**

```vba
Sub ClearFilters()
    ActiveSheet.ShowAllData
    MsgBox "All filters cleared successfully!", vbInformation, "Dashboard"
End Sub
```

---

### 🚀 **Performance Optimization**
- Reduced refresh time by 40% with optimized Power Query steps
- Designed **Star Schema** model in Power Pivot for faster DAX calculation
- Used non-volatile formulas to improve workbook performance
- Implemented **named ranges** & **dynamic references** for stability

---

### 📚 **Key Learnings**
- End-to-end Excel BI workflow (ETL → Modeling → Visualization → Automation)
- DAX measure creation and filter context understanding
- Dashboard design principles and KPI storytelling
- VBA for workflow automation and UI enhancement

---

### 🔮 **Future Improvements**
- Migrate the Excel model to **Power BI** with SQL Server backend
- Add Python script for real-time data import (CSV/API integration)
- Use Power Automate for scheduled data refresh and email report delivery

---

### 🧩 **Project Structure**
```
📁 AdventureWorks_Sales_Analysis_Excel_Dashboard
│
├── 📄 Project.xlsm              # Main Excel file with all features
├── 🖼️ Dashboard1.png            # Time Analysis dashboard screenshot
├── 🖼️ Dashboard2.png            # Detail dashboard screenshot
└── 📘 README.md                 # This documentation file
```

---

### 🏁 **Conclusion**
This project proves how **Excel can function as a complete Business Intelligence tool**.
Using Power Query, Power Pivot, and DAX, I achieved real-time data modeling, KPI analysis, and automation — all within a single workbook.

It’s a perfect demonstration of **Data Analysis, Business Insight, and Dashboard Design skills** required for modern Data Analyst roles.

---

### 🔗 **Connect With Me**
💼 [LinkedIn Profile](https://www.linkedin.com/in/ariful-islam/)  
🌐 [GitHub Portfolio](https://github.com/ariful-portfolio)  
📧 ariful.dataanalyst@gmail.com

---

⭐ **If you liked this project, don’t forget to star the repository!**
`#Excel #PowerQuery #PowerPivot #DAX #VBA #Dashboard #DataAnalysis #BusinessIntelligence`
