# 📊 GL Insights  

**GL Insights** is a free financial analytics tool built for **general ledger analysis**. Upload your general ledger Excel file, and the program will generate insightful datasets, detect errors, and highlight potential fraud risks—helping accountants and businesses make data-driven decisions efficiently.  

## 🎯 Features  
✅ **Trend Analysis for Every Account** – Identify seasonal trends, strengths, and weaknesses using historical data. Implements **ETS (Exponential Triple Smoothing)** for forecasting and anomaly detection.  
✅ **Error Detection** – Automatically flags unbalanced and high-variance transactions within each account.  
✅ **Fraud Alert** – Detects suspicious financial activities, such as **account clearing** and **transaction offsetting**.  
✅ **Financial Performance Insights** – Transforms raw ledger data into meaningful narratives for better decision-making.  

## 🛠️ Technologies Used  
- **Programming Language:** Python, VBA
- **Libraries:** Openpyxl

## 📖 How to Use  
1. Open the **GL Insights** executable file.  
2. Select a **general ledger (GL) Excel file** in **QuickBooks format**.  

### 📂 GL File Format Requirements  
- Must contain **only one tab**, named **"GL"**.  
- Must have the following **columns (not below row 10)**:  
  - **Date** (formatted as `YYYY-MM-DD`, type: "Date")  
  - **Num**, **Split**, **Debit**, **Credit**, **Balance**  
- **Column B** must define accounts, numbered properly:  
  - **1xxx / 1xxxx** → Assets  
  - **2xxx / 2xxxx** → Liabilities  
  - **3xxx / 3xxxx** → Equity  
  - **4xxx / 4xxxx** → Revenue  
  - **5xxx - 9xxx / 5xxxx - 9xxxx** → Expenses / COGS  
- Ensure **"gl-insights-trends.xlsm"** is in the same folder for trend analysis.  

GL Insights will then analyze the data, detect anomalies, and provide insights automatically. 

## 🚀 Lessons Learned  
Developing **GL Insights** taught me how to automate financial analysis by detecting:  
- **Unusual expenses & high-variance entries**  
- **Account clearing & transaction offsetting**  
- **Reclassifications & unbalanced journal entries**  

This project reinforced my understanding of **accounting best practices, error detection, and fraud prevention**, while improving my ability to **analyze and automate financial workflows**.  

## 🖼️ Screenshots  
<p align="center">
  <img src="https://github.com/user-attachments/assets/1c115c1e-a018-4268-813a-6f9d8b01a65a" width="60%" />
  <img src="https://github.com/user-attachments/assets/8e6e47ba-6fae-4106-b0be-b4f73997b1a4" width="35%"/>
</p>

---
