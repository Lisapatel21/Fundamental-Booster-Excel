[README (1).md](https://github.com/user-attachments/files/33827669/README.1.md)
# PR. 1 Fundamental Booster

**Excel / Google Sheets Formula Project: Students, Sales and Employee Analytics**

> Workbook: `_PR__1_Fundamental_Booster.xlsx` | Sheets: 4 | Records: 20 per data sheet
> Values quoted in this guide were computed as of **22 Oct 2026** (date-dependent results are marked 📅).

---

## 📑 Table of Contents

1. [Project Overview](#1--project-overview)
2. [Workbook Structure](#2--workbook-structure)
3. [Audit: Issues Found in the Current File](#3--audit-issues-found-in-the-current-file-fix-before-submitting)
4. [Sheet 1: Students Grade](#4--sheet-1-students-grade)
5. [Sheet 2: Sales Data](#5--sheet-2-sales-data)
6. [Sheet 3: Employee Data](#6--sheet-3-employee-data)
7. [Reference Rules: Where to Lock with `$`](#7--reference-rules-where-to-lock-with-)
8. [Step-by-Step Submission Guide](#8--step-by-step-submission-guide)
9. [Submission Checklist](#9--submission-checklist)
10. [Troubleshooting Formula Errors](#10--troubleshooting-formula-errors)
11. [Excel vs Google Sheets Compatibility](#11--excel-vs-google-sheets-compatibility)

---

## 1. 🎯 Project Overview

This project practises core spreadsheet analytics on three datasets.

| Topic | Functions Covered |
|---|---|
| Conditional logic | `IF`, nested `IF`, `IF` + `AND` / `OR` |
| Conditional statistics | `COUNTIFS`, `SUMIFS`, `AVERAGEIFS` |
| Lookups | `VLOOKUP`, `XLOOKUP`, `INDEX` + `MATCH`, `XMATCH` |
| Text | `LEFT`, `FIND`, `TRIM`, `UPPER`, `LOWER` |
| Dates | `TODAY`, `DATEDIF`, date subtraction |
| Math | `ROUND`, `CEILING`, `FLOOR` |
| Dynamic references | `INDIRECT`, `OFFSET`, `FILTER` |
| References | Relative (`A1`), absolute (`$A$1`), mixed (`$A1`, `A$1`) |

**Deliverable:** complete each worksheet by applying the required formulas in the right sheet.

---

## 2. 🗂️ Workbook Structure

| # | Sheet | Data Range | Input Columns | Formula Columns |
|---|---|---|---|---|
| 0 | Project Instructions | `A1:A56` | n/a | n/a (read only) |
| 1 | Students Grade | `A1:L21` | A–F | G–L |
| 2 | Sales Data | `A1:M21` | A–F, J | G–I, K–M |
| 3 | Employee Data | `A1:L21` | A–E | F–L |

### Data dictionary

| Sheet | Columns |
|---|---|
| **Students Grade** | Student ID · Name · Math · Science · English · Enrollment Date |
| **Sales Data** | Sales ID · Product · Region · Salesperson · Amount · Date |
| **Employee Data** | Employee ID · Name · Department · Salary · Joining Date |

---

## 3. 🔍 Audit: Issues Found in the Current File (Fix Before Submitting)

I read every formula in your workbook. Most are sound, but these issues would cost marks or produce wrong answers.

| # | Sheet / Cells | Issue | Severity | Fix |
|---|---|---|---|---|
| 1 | Students `J2:L21` | **Ranges drift downward.** `J2` uses `C2:C21`, but `J3` uses `C3:C22`, and `J21` uses `C21:C40`. The count shrinks row by row (19, 18, 17 … 1). Only `J2` is correct. | 🔴 High | Lock the range: `$C$2:$C$21` (see §4) |
| 2 | Students `K2:K21` | Header says *"Total Marks for High Performers across all subjects"*, but the formula sums the **Average** column (`G`) filtered on Math > 80. It returns 753.33, not total marks. | 🔴 High | Sum Math + Science + English for `"High Scorer"` rows (see §4). Correct answer: **768** |
| 3 | Employee `I19:I21` | `DATEDIF` returns **`#NUM!`** (the start date is later than today). Joining dates `2027-04-28` and `2028-01-03` are in the future. 📅 | 🔴 High | Guard with `IF(E2>TODAY(), …)` (see §6) and confirm those dates with your instructor |
| 4 | Sales `G` vs `H` | Boundary mismatch: `G` treats `>=30000` as the top tier, but `H` uses `>30000`. A sale of exactly 30,000 gets the top discount but "Not Eligible". | 🟡 Medium | Use `>=30000` in both |
| 5 | Sales `I2:I21` | `"South"` and `"Keyboard"` are **hard-coded** inside the formula and repeated 20 times. | 🟡 Medium | Move to input cells `O2` / `P2`, keep one formula |
| 6 | Instructions sheet | Conflict: Task 2 says Math **above 60**; the Essentials list says **above 50**. Both give 19 with your data, but state your choice. | 🟡 Medium | Keep `>60` (matches the sheet header). Note it in your submission |
| 7 | Employee `B` | Trailing spaces in `Aarav Sharma␣`, `Priya Verma␣`, `Arjun Singh␣`. `UPPER` output keeps the space. | 🟢 Low | Wrap in `TRIM()` |
| 8 | Employee `G` | `XLOOKUP(A2, A:A, D:D)` looks up each row's own ID, so it always succeeds and demonstrates nothing. | 🟢 Low | Drive it from an input cell (`P2`, see §6) |
| 9 | Students `F19:F20` | Enrollment dates `2026-11-15` and `2027-05-14` are in the future. 📅 | 🟢 Low | Verify the data; no formula uses them yet |
| 10 | Whole file | **Missing deliverables:** `FILTER`, `LOWER`, `INDIRECT`, `OFFSET`, student-name `VLOOKUP`, salesperson `XLOOKUP`, day-difference calculation. | 🔴 High | Added in §4–§6 |

---

## 4. 🎓 Sheet 1: Students Grade

### 4.1 Per-row formulas (fill down rows 2–21)

| Cell | Purpose | Formula |
|---|---|---|
| `G2` | Average of three subjects | `=AVERAGE(C2:E2)` |
| `H2` | Letter grade (**nested IF**) | `=IF(G2>=90,"A",IF(G2>=80,"B",IF(G2>=70,"C",IF(G2>=60,"D","F"))))` |
| `I2` | Above 80 in Math **and** Science (**IF + AND**) | `=IF(AND(C2>80,D2>80),"High Scorer","Regular")` |

> `C2:E2` is intentionally **relative**: each row must average its own marks.

### 4.2 Summary formulas (enter **once**, in row 2 only, then clear `J3:L21`)

| Cell | Purpose | Formula | Expected |
|---|---|---|---|
| `J2` | Students scoring above 60 in Math (**COUNTIFS**) | `=COUNTIFS($C$2:$C$21,">60")` | **19** |
| `K2` | Total marks of high performers, all subjects (**SUMIFS**) | `=SUMIFS($C$2:$C$21,$I$2:$I$21,"High Scorer")+SUMIFS($D$2:$D$21,$I$2:$I$21,"High Scorer")+SUMIFS($E$2:$E$21,$I$2:$I$21,"High Scorer")` | **768** |
| `L2` | Average Math score for students above 60 (**AVERAGEIFS**) | `=AVERAGEIFS($C$2:$C$21,$C$2:$C$21,">60")` | **82.26** |

High scorers are Student 12 (253), Student 13 (285) and Student 20 (230), which total 768.

### 4.3 Added deliverables (new cells in columns N–U)

Label the input and output cells in row 1, as shown.

| Cell | Label | Formula / Value | Expected |
|---|---|---|---|
| `N1` | *Lookup Student ID* | (label) | |
| `N2` | Input | `1005` (type the value) | |
| `O1` | *Student Name (VLOOKUP)* | (label) | |
| `O2` | **VLOOKUP**: name by ID | `=VLOOKUP($N$2,$A$2:$B$21,2,FALSE)` | Student 5 |
| `N4` | *FILTER: Average > 80* | (label) | |
| `N5` | **FILTER** (spills to `N5:U6`) | `=FILTER($A$2:$H$21,$G$2:$G$21>80,"No students found")` | 2 rows: Students 12 and 13 |

> **FILTER spill rule:** cells `N5:U…` below and to the right must be **empty**, or Excel shows `#SPILL!`.
> For a "top performers" list sorted high to low, use `=SORT(FILTER($A$2:$H$21,$G$2:$G$21>80,"None"),7,-1)`.

---

## 5. 💰 Sheet 2: Sales Data

### 5.1 Per-row formulas (fill down rows 2–21)

| Cell | Purpose | Formula |
|---|---|---|
| `G2` | Discount amount (**nested IF**): 15% ≥ 30,000 · 10% ≥ 15,000 · else 5% | `=IF(E2>=30000,E2*0.15,IF(E2>=15000,E2*0.1,E2*0.05))` |
| `H2` | Discount eligibility (**IF + OR**): Laptop *or* high-value sale | `=IF(OR(B2="Laptop",E2>=30000),"Eligible","Not Eligible")` |

> ✏️ `H` is corrected from `>30000` to `>=30000` so it matches `G`.

### 5.2 Input cells (type these values)

| Cell | Label (row 1) | Value |
|---|---|---|
| `O2` | Region | `South` |
| `P2` | Product | `Keyboard` |
| `Q2` | Salesperson | `Person 10` |
| `R2` | Month Start | `2023-05-01` (formatted as a date) |
| `J2` | Search Input (Sales ID) | `2001` *(one cell only; clear `J3:J21`)* |
| `U2` | Last Row (INDIRECT) | `11` |
| `W2` | N Latest (OFFSET) | `5` |

### 5.3 Summary and lookup formulas (enter **once**, in row 2; clear rows 3–21 of these columns)

| Cell | Task | Formula | Expected |
|---|---|---|---|
| `I2` | **SUMIFS**: sales for region + product | `=SUMIFS($E$2:$E$21,$C$2:$C$21,$O$2,$B$2:$B$21,$P$2)` | **42,233** |
| `K2` | **VLOOKUP**: product by Sales ID | `=VLOOKUP($J$2,$A$2:$E$21,2,FALSE)` | Keyboard |
| `L2` | **INDEX/MATCH**: amount by Sales ID | `=INDEX($E$2:$E$21,MATCH($J$2,$A$2:$A$21,0))` | 7,048 |
| `M2` | **XMATCH**: position of the ID | `=XMATCH($J$2,$A$2:$A$21)` | 1 |
| `S2` | **INDEX/MATCH** (two criteria): salesperson + month | `=INDEX($E$2:$E$21,MATCH(1,($D$2:$D$21=$Q$2)*(MONTH($F$2:$F$21)=MONTH($R$2))*(YEAR($F$2:$F$21)=YEAR($R$2)),0))` | 35,185 |
| `T2` | **XLOOKUP**: salesperson's amount | `=XLOOKUP($Q$2,$D$2:$D$21,$E$2:$E$21,"Not Found")` | 35,185 |
| `V2` | **INDIRECT**: dynamic range | `=SUM(INDIRECT("E2:E"&$U$2))` | 249,975 |
| `X2` | **OFFSET**: sum of the latest N sales | `=SUM(OFFSET($E$2,COUNT($E$2:$E$1000)-$W$2,0,$W$2,1))` | 149,698 |
| `Y2` | **Date difference** in days | `=MAX($F$2:$F$21)-MIN($F$2:$F$21)` | 285 |

> **Your existing robust variants also work** (they handle IDs stored as text):
> `=IFNA(VLOOKUP(VALUE($J$2),$A$2:$E$21,2,FALSE),VLOOKUP(TEXT($J$2,"#"),$A$2:$E$21,2,FALSE))`
> Note `J2` must be **locked** (`$J$2`) if you keep the formula in several rows.

> **`S2` is an array formula.** It works as-is in Excel 365 and Google Sheets. In Excel 2016/2019, confirm with **Ctrl + Shift + Enter**.

---

## 6. 👥 Sheet 3: Employee Data

### 6.1 Per-row formulas (fill down rows 2–21)

| Cell | Task | Formula |
|---|---|---|
| `F2` | **First name** (LEFT + FIND, space-safe) | `=LEFT(TRIM(B2),FIND(" ",TRIM(B2)&" ")-1)` |
| `H2` | **UPPER** | `=UPPER(TRIM(B2))` |
| `M2` | **LOWER** *(new column, header "Convert Name To Lowercase")* | `=LOWER(TRIM(B2))` |
| `I2` | **Years of service** (DATEDIF + TODAY, future-safe) | `=IF(E2>TODAY(),"Not joined yet",DATEDIF(E2,TODAY(),"Y"))` |
| `N2` | **Days since joining** *(new, header "Days Since Joining")* | `=IF(E2>TODAY(),"Not joined yet",TODAY()-E2)` |
| `J2` | **ROUND** 10% bonus to whole number | `=ROUND(D2*0.1,0)` |
| `K2` | **CEILING** 10% bonus up to the next 100 | `=CEILING(D2*0.1,100)` |
| `L2` | **FLOOR** 10% bonus down to the previous 100 | `=FLOOR(D2*0.1,100)` |

> `FIND(" ",TRIM(B2)&" ")` appends a space so a single-word name cannot cause `#VALUE!`.

### 6.2 Lookup panel (type the inputs, enter formulas once)

| Cell | Label (row 1) | Content | Expected |
|---|---|---|---|
| `P2` | Lookup Employee ID | `3005` | |
| `Q2` | Lookup Field | `Department` | |
| `R2` | **XLOOKUP** Salary | `=XLOOKUP($P$2,$A$2:$A$21,$D$2:$D$21,"Not Found")` | 54,149 |
| `S2` | **INDEX/MATCH** any field | `=INDEX($A$2:$E$21,MATCH($P$2,$A$2:$A$21,0),MATCH($Q$2,$A$1:$E$1,0))` | IT |

`S2` returns whichever column `Q2` names, with no column-position restriction. Try `Salary`, `Name` or `Joining Date`.

### 6.3 Sample results 📅 (as of 22 Oct 2026)

| Employee | Salary | `ROUND` | `CEILING` | `FLOOR` | Years |
|---|---|---|---|---|---|
| Aarav Sharma | 86,766 | 8,677 | 8,700 | 8,600 | 11 |
| Priya Verma | 109,972 | 10,997 | 11,000 | 10,900 | 9 |
| Davendra Pandey | 116,283 | 11,628 | 11,700 | 11,600 | *Not joined yet* |

---

## 7. 🔒 Reference Rules: Where to Lock with `$`

| Situation | Use | Example | Why |
|---|---|---|---|
| Lookup table / criteria range, copied down | **Absolute** `$A$2:$E$21` | `VLOOKUP($J$2,$A$2:$E$21,2,FALSE)` | Prevents the range sliding (the cause of `#N/A` and wrong counts) |
| Single input cell reused by many rows | **Absolute** `$J$2` | `MATCH($J$2,…)` | Every row must read the same input |
| Value that must change per row | **Relative** `E2` | `IF(E2>=30000,…)` | Each row evaluates its own data |
| Column fixed, row moves | **Mixed** `$A2` | `$A2*B$1` | Multiplication-table layouts |
| Header row used as lookup row | **Mixed** `$A$1:$E$1` | `MATCH($Q$2,$A$1:$E$1,0)` | Prevents header reference drift |

**Shortcut:** select the reference inside the formula bar and press **F4** (Windows) or **⌘ + T** (Mac) to cycle `A1 → $A$1 → A$1 → $A1`.

---

## 8. 📝 Step-by-Step Submission Guide

### Step 1: Back up
Save a copy: `PR1_Fundamental_Booster_BACKUP.xlsx`.

### Step 2: Fix the audit items (§3)
1. Students: delete `J3:L21`, then re-enter `J2`, `K2`, `L2` from §4.2.
2. Sales: change `H2` to `>=30000` and fill down; delete `I3:I21`; keep `I2` from §5.3.
3. Employee: replace `F`, `H`, `I` using §6.1.

### Step 3: Add the missing deliverables
1. **Students:** `N1:O2` (VLOOKUP) and `N4:N5` (FILTER). Keep `N5:U…` clear.
2. **Sales:** input cells (§5.2), then `S2`, `T2`, `V2`, `X2`, `Y2`.
3. **Employee:** columns `M`, `N`, then the lookup panel `P1:S2`.

### Step 4: Format
- Bold header rows; freeze row 1 (**View → Freeze Panes**).
- Dates: `dd-mmm-yyyy` · Currency (Amount, Salary): `#,##0` or `₹#,##0`.
- Averages: 2 decimals (`0.00`).
- Shade input cells light yellow and output cells light green so reviewers can see which is which.

### Step 5: Verify against expected values

| Check | Cell | Expected |
|---|---|---|
| Math > 60 count | Students `J2` | 19 |
| High-performer total | Students `K2` | 768 |
| Avg Math > 60 | Students `L2` | 82.26 |
| FILTER rows | Students `N5` | Students 12 and 13 |
| South + Keyboard sales | Sales `I2` | 42,233 |
| Eligible count (`H2:H21`) | `=COUNTIF($H$2:$H$21,"Eligible")` | 13 |
| INDEX/MATCH | Sales `L2` | 7,048 |
| Employee 3005 salary | Employee `R2` | 54,149 |
| Errors in workbook | Press **Ctrl + F** and search `#` | None |

### Step 6: Final error sweep
Run **Home → Find & Select → Go To Special → Formulas → Errors** on each sheet. The selection must be empty.

### Step 7: Add a short notes block
On the Instructions sheet (or a new `Notes` sheet), record:
- The threshold chosen for Task 2 (`>60`).
- That future dates exist in `Employee!E19:E20` and `Students!F19:F20`.
- That `FILTER` uses the **Average** column (`>80`).

### Step 8: Save and submit
Save as `PR1_Fundamental_Booster_<YourName>.xlsx`. If submitting via Google Sheets, use **Share → Anyone with the link → Viewer**, and also export an `.xlsx`.

---

## 9. ✅ Submission Checklist

| ☐ | Requirement | Where |
|---|---|---|
| ☐ | Nested IF grading | Students `H` |
| ☐ | IF + AND (Math and Science > 80) | Students `I` |
| ☐ | Nested IF discount | Sales `G` |
| ☐ | IF + OR eligibility | Sales `H` |
| ☐ | COUNTIFS / SUMIFS / AVERAGEIFS with `$` ranges | Students `J2:L2`; Sales `I2` |
| ☐ | VLOOKUP (student name; product) | Students `O2`; Sales `K2` |
| ☐ | INDEX/MATCH (single and two-way) | Sales `L2`, `S2`; Employee `S2` |
| ☐ | XLOOKUP | Sales `T2`; Employee `R2` |
| ☐ | XMATCH | Sales `M2` |
| ☐ | LEFT + FIND, UPPER, LOWER | Employee `F`, `H`, `M` |
| ☐ | DATEDIF + TODAY (no `#NUM!`) | Employee `I` |
| ☐ | Date difference in days | Employee `N`; Sales `Y2` |
| ☐ | ROUND / CEILING / FLOOR | Employee `J`, `K`, `L` |
| ☐ | INDIRECT and OFFSET | Sales `V2`, `X2` |
| ☐ | FILTER | Students `N5` |
| ☐ | No error values anywhere | All sheets |
| ☐ | Formatting applied, header row frozen | All sheets |

---

## 10. 🛠️ Troubleshooting Formula Errors

| Error | Typical cause in this project | Fix |
|---|---|---|
| **`#N/A`** | Lookup value not found: the ID is stored as **text** in one place and a **number** in the other; or the lookup range slid down because `$` is missing | Use `$A$2:$E$21`; convert with `VALUE()` / `TEXT(J2,"#")`; wrap in `IFNA(…,"Not Found")` |
| **`#REF!`** | A referenced column or sheet was deleted, or `VLOOKUP`'s column index exceeds the table width (e.g. `6` on `$A$2:$E$21`) | Re-check the index (max `5` here); use `INDEX/MATCH` so column deletions don't break it |
| **`#NAME?`** | Typo in a function name, or `XLOOKUP`/`XMATCH`/`FILTER` in Excel 2016/2019 | Check spelling; use Microsoft 365 / Excel 2021+, or fall back to `INDEX/MATCH` |
| **`#VALUE!`** | `FIND` can't find a space (single-word name); text used in arithmetic; wrong data type | Use `FIND(" ",TRIM(B2)&" ")`; confirm numbers aren't text |
| **`#NUM!`** | `DATEDIF` with a **start date later than today**, as in `E19:E20` | `IF(E2>TODAY(),"Not joined yet",DATEDIF(E2,TODAY(),"Y"))` |
| **`#DIV/0!`** | `AVERAGEIFS` with no matching rows | `=IFERROR(AVERAGEIFS(…),0)` |
| **`#SPILL!`** | `FILTER` output blocked by data in the spill range | Clear `N5:U…`; avoid merged cells in that area |
| **`#CALC!`** | `FILTER` found no matches and no "if empty" argument | Add the 3rd argument: `FILTER(…,"No students found")` |
| **`#NULL!`** | Space used instead of `:` or `,` in a range | `A2:A21`, not `A2 A21` |
| **Numbers don't update** | Calculation set to Manual | **Formulas → Calculation Options → Automatic** |
| **Count shrinks row by row** | Relative range copied down (**the bug in `J2:L21`**) | Lock with `$C$2:$C$21` |
| **Shows formula text, not result** | Cell formatted as Text | Set to **General**, re-enter with **F2 → Enter** |
| **Wrong `UPPER`/`LEFT` output** | Hidden leading/trailing spaces | Wrap with `TRIM()` |
| **`DATEDIF` not in the function list** | It's a "hidden" function in Excel | Type it manually; it still works |
| **Dates show as numbers (e.g. 45123)** | Cell format is General | Set the format to **Short Date** |

### Debugging workflow
1. Select the broken cell → **Formulas → Evaluate Formula** to step through it.
2. Use **Formulas → Trace Precedents** to see which ranges it reads.
3. Test pieces in a scratch cell: `=MATCH($J$2,$A$2:$A$21,0)`.
4. Check data types with `=ISNUMBER(A2)` / `=ISTEXT(A2)`.

---

## 11. 🔄 Excel vs Google Sheets Compatibility

| Function | Excel 365 / 2021 | Excel 2016/2019 | Google Sheets |
|---|---|---|---|
| `IF`, `COUNTIFS`, `SUMIFS`, `AVERAGEIFS`, `VLOOKUP`, `INDEX/MATCH` | ✅ | ✅ | ✅ |
| `XLOOKUP`, `XMATCH`, `FILTER` | ✅ | ❌ | ✅ (verify `XMATCH`; fall back to `MATCH`) |
| `DATEDIF`, `TODAY` | ✅ | ✅ | ✅ |
| `INDIRECT`, `OFFSET` | ✅ | ✅ | ✅ |
| Two-criteria `INDEX/MATCH` (`S2`) | ✅ native | ⚠️ Ctrl + Shift + Enter | ✅ wrap in `ARRAYFORMULA(...)` if it returns `#N/A` |

**Opening the `.xlsx` in Google Sheets:** formulas convert automatically, but re-check `XMATCH`, `FILTER` spills and `TODAY()`-based results.

---

### 📌 Quick Reference: Core Patterns

```excel
Nested IF      =IF(G2>=90,"A",IF(G2>=80,"B",IF(G2>=70,"C",IF(G2>=60,"D","F"))))
IF + AND/OR    =IF(AND(C2>80,D2>80),"High Scorer","Regular")
COUNTIFS       =COUNTIFS($C$2:$C$21,">60")
SUMIFS         =SUMIFS($E$2:$E$21,$C$2:$C$21,$O$2,$B$2:$B$21,$P$2)
AVERAGEIFS     =AVERAGEIFS($C$2:$C$21,$C$2:$C$21,">60")
VLOOKUP        =VLOOKUP($N$2,$A$2:$B$21,2,FALSE)
INDEX/MATCH    =INDEX($E$2:$E$21,MATCH($J$2,$A$2:$A$21,0))
XLOOKUP        =XLOOKUP($P$2,$A$2:$A$21,$D$2:$D$21,"Not Found")
XMATCH         =XMATCH($J$2,$A$2:$A$21)
First name     =LEFT(TRIM(B2),FIND(" ",TRIM(B2)&" ")-1)
DATEDIF        =IF(E2>TODAY(),"Not joined yet",DATEDIF(E2,TODAY(),"Y"))
FILTER         =FILTER($A$2:$H$21,$G$2:$G$21>80,"No students found")
```

---

**Author:** Lisa Patel

⭐ Thank you for reading! If this guide helped you build a clean, error-free submission, a star on the project would mean a lot. Happy analysing! 🙏
