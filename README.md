# Investment Scenario Analysis Tool (VBA)

This repository contains all components of my **MANG2092 – Business Analytics Programming** coursework submission.
The tool models investment outcomes for **SW Asset Management’s £10 billion portfolio**, helping the user test different economic scenarios.

---

## 📁 Files Included

### **1. Ques_VBa.pdf**

The coursework (MANG2092 – Business Analytics Programming) requires building a VBA program to support SW Asset Management’s CFO, Terry, in analysing long-term investment scenarios for a £10 billion fund. The brief outlines three investment classes, each with fixed sub-sector allocation rules, including areas such as AI & IT, Banks, Energy, Global Supply Chain, Real Estate, and Smart Automotive. Each sub-sector has an annual return range that the user can explore.

The assignment expects the tool to:

- Allow users to enter expected annual return rates for each sub-sector.

- Accept an investment duration (e.g., 1, 3, or 5 years).

- Validate data inputs, ensuring values are non-negative and percentages follow logical rules.

- Display errors using MsgBox where necessary.

- Calculate sub-sector, class-level, and total portfolio returns using compound interest.

- Present results clearly to the user.

The brief also encourages extensions that improve flexibility and usability. Suggested enhancements include looping multiple scenarios, making the function available directly in Excel (Public Function), creating better data entry interfaces, and allowing different input formats. A reflective paragraph explaining the chosen extensions is required as part of the submission.

Additionally, the coursework includes strict formatting, submission, and academic integrity requirements, along with instructions to submit both a PDF of the VBA routine and an xlsm file containing the working model.

### **2. PDF_extension.docx**

Full VBA routine including:

* Workbook_Open welcome message
* UserForm-based input interface
* Scenario auto-fill (worst / best case)
* Validation and error handling
* Compound interest return calculations
* Logging to *Investment_Log* sheet
* Public function `CalculateReturn`
* Storyline and extension explanation


### **3. Excel Macro-Enabled Workbook (.xlsm)**

Implements the full investment calculator with:

* UserForm for data entry
* Automated log sheet
* Scenario testing loop
* Callable public function for worksheets

---

## ⚙️ How the Tool Works

The tool calculates returns for three investment classes based on:

* User-entered class allocations
* Sub-sector return rates
* Investment duration
* Optional auto-fill using minimum/maximum returns from the brief

It computes **compound returns** using:

```
Return = PV × (1 + r)^t – PV
```

Results are shown via message box and stored in **Investment_Log** for tracking multiple tests.

---

## 🚀 Features

* Clean UserForm interface
* Automated validation of all inputs
* Worst / Best case auto-fill
* Multi-scenario loop
* Public worksheet function
* Timestamped logging

---

## 🛠️ Requirements

* Microsoft Excel with macros enabled
* VBA runtime (built into Excel)

---

## 🔧 How to Clone the Repository & Run the Excel Tool

### **1️⃣ Clone the Repository**

You can clone the repo using HTTPS or SSH.

**Using HTTPS:**

```bash
git clone https://github.com/<your-username>/<your-repo-name>.git
```

**Using SSH:**

```bash
git clone git@github.com:<your-username>/<your-repo-name>.git
```

Then open the folder:

```bash
cd <your-repo-name>
```

---

## **2️⃣ Open the Excel File (.xlsm)**

1. Locate the file:
   **Excel_VBA.xlsm**
   (This is your macro-enabled workbook containing the full UserForm and VBA logic.)

2. Double-click the file to open it in **Microsoft Excel**.

---

## **3️⃣ Enable Macros (Important)**

Because the tool uses VBA, Excel will block macros by default.
Enable them using:

**When Excel opens:**

* Click **“Enable Content”** on the yellow security banner
  *or*
* Go to **File → Info → Enable Content → Enable All Macros**

---

## **4️⃣ Run the Scenario Calculator**

Once macros are enabled:

### Option A — Use the Dashboard Button

If your workbook includes a dashboard button:

* Click **“Start Scenario Calculator”**
  This will open the UserForm.

### Option B — Run from Developer Tab

1. Go to **Developer → Macros**
2. Select **OpenScenarioCalculator**
3. Click **Run**

This opens the **Return Calculation Form**, where you can enter:

* Class allocation percentages
* Sub-sector return rates
* Investment duration (years)
* Optional Worst/Best Case Auto-Fill

Click **Calculate** to:

* Display detailed returns
* Automatically log results in the *Investment_Log* sheet
* Choose whether to run another scenario

---

## **5️⃣ Optional: Use the Public Function in Excel**

You can also calculate return directly in a worksheet using:

```excel
=CalculateReturn(amount, rate, years)
```

Example:

```excel
=CalculateReturn(2000000000, 8, 5)
```
## 📜 Licence

Personal academic project. Not intended for commercial use.

