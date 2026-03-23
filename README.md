# 📈 Equity Research & Valuation Engine

### Automated Fundamental Analysis, Forecasting & Intrinsic Valuation (Excel + Python + VBA)

A hybrid financial modeling system that automates **data extraction, financial analysis, forward projections, and valuation** — all within Excel.

---

## 🚀 What This Does

This engine allows you to:

* Input a **NASDAQ ticker**
* Automatically fetch **5 years of financial data**
* Build structured **3 financial statements**
* Analyze **historical performance (scored)**
* Generate **5-year forward projections**
* Estimate **intrinsic value (DCF + multiples)**
* Get a **final recommendation with guardrails**

---

## 🧠 Model Architecture

The model is built in four logical layers:

### 1. Historical Analysis

* Growth, profitability, cash flow quality
* Financial strength & capital allocation
* Scoring system to standardize evaluation

### 2. Forward Projections (5 Years)

* Revenue growth based on assumptions
* Margin evolution & reinvestment logic
* Working capital modeling

### 3. Valuation Engine

* Discounted Cash Flow (DCF)
* Multiple-based valuation
* Implied market expectations
* Return decomposition

### 4. Decision Layer

* Final score & recommendation
* Valuation regime (premium/discount)
* Guardrails to flag risky assumptions

---

## ⚙️ Setup (2–3 Minutes)

### 1. Install Dependencies

```bash
pip install -r requirements.txt
```

### 2. Add API Key

Create a file:

```
api_key.py
```

Add:

```python
API_KEY = "your_fmp_api_key"
```

### 3. Enable Excel Macros

* Open `Equity_Analysis_Engine.xlsm`
* Enable macros when prompted

---

## ▶️ How to Use

1. Open the Excel model
2. Enter a **NASDAQ ticker**
3. Click **Fetch**
4. Wait ~5–10 seconds

That’s it — the model will:

* Pull data via Python
* Build statements
* Run analysis
* Output valuation & recommendation

---

## 🎯 Manual Inputs

Some assumptions are user-controlled (highlighted in **blue** in Excel):

* Industry growth rate
* Capex intensity
* Risk-free rate, ERP, Beta (WACC inputs)
* WACC adjustment
* Terminal growth rate (g)

⚠️ Small changes here can significantly impact valuation.

---

## 🛠️ Tech Stack

* **Python** → Data pipeline (API + processing)
* **pandas / requests / xlwings** → Data handling & Excel integration
* **VBA** → Automation trigger layer
* **Excel** → Core modeling & dashboard

---

## 💡 Why This Project Stands Out

* Not just analysis — **end-to-end system**
* Combines **finance + programming + automation**
* Handles **real-world messy financial data**
* Implements **forward-looking valuation logic**
* Built like a **mini equity research engine**

---

## ⚠️ Limitations

* Works with **NASDAQ-listed tickers**
* Dependent on API availability/limits
* Requires Excel macros enabled

---

## 🔮 Future Improvements

* Multi-company comparison
* Scenario & sensitivity dashboard
* Portfolio tracking
* Web-based interface (Streamlit)

---

⭐ If you found this useful, consider starring the repo!

