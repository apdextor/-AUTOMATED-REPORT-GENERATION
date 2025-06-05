# AUTOMATED REPORT GENERATION

**Company:** Code Tech IT Solutions  
**Name:** Ayush Patil  
**Intern ID:** CT06DN877  
**Domain:** Python Development  
**Duration:** 6 weeks  
**Mentor:** Neela Santosh  


![Report Banner](https://img.shields.io/badge/PDF-Report-blue?style=flat-square&logo=adobeacrobatreader)

## 📊 Project Overview

This project is a **comprehensive sales analytics and reporting tool** built with Python. It generates a visually appealing, data-driven PDF report from sales data, providing insights into sales performance by category, country, product, and month.

- **Input:** Sales data (CSV)
- **Output:** Professional PDF report with charts, tables, and executive summary

---

## ✨ Features

- **Automatic Data Generation:** Creates realistic sample sales data for demonstration.
- **In-depth Analysis:** Calculates total sales, average sale, top categories, products, and regions.
- **Beautiful PDF Report:** Uses [ReportLab](https://www.reportlab.com/) to generate a multi-page, styled PDF with:
  - Cover page & executive summary
  - Bar, pie, and line charts for categories, products, countries, and months
  - Key metrics and tables
  - Recommendations section

---

## 🚀 How to Use

1. **Install Requirements**

   ```sh
   pip install pandas numpy reportlab
   ```

2. **Run the Script**

   ```sh
   python task2.py
   ```

   - This will generate:
     - `sales_data.csv` (sample data)
     - `Sales_Report_2024.pdf` (the final report)

3. **View the Report**

   - Open `Sales_Report_2024.pdf` in any PDF viewer.

---

## 📎 Example Output

> **Preview:**  
> ![PDF Preview](https://user-images.githubusercontent.com/placeholder/pdf-preview.png)  
> *(Replace with a screenshot of your PDF report)*

**[Download the latest report &raquo;](Sales_Report_2024.pdf)**

---

## 📁 Project Structure

```
├── sales_data.csv
├── Sales_Report_2024.pdf
└── task2.py
```

---

## 🛠️ Customization

- To use your own data, replace `sales_data.csv` with your dataset (ensure columns match).
- Adjust the number of rows or categories in `task2.py` as needed.

---

## 📄 License

MIT License

---

> _Generated with ❤️ using Python, Pandas, and ReportLab._
