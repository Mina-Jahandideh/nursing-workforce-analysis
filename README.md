# Employment Status Change Analysis (VBA)

This repository contains VBA macros for analyzing employment data related to nurses and midwives. The project focuses on tracking employment movements, status changes, and terminations, with specific attention to roles such as GRN (Graduated Registered Nurse) and GRM (Graduated Registered Midwife).

⚠️ **Note:** Due to privacy concerns, the actual dataset is not included.

---

## 📦 Contents

- VBA macros to:
  - Count employment status changes
  - Identify terminated subjects
  - Track position movements
  - Extract last known status
  - Tag GRNs and GRMs

---

## 🧾 Data Requirements

The macros are designed to work with an Excel file containing the following sheets:

- **`Position Movements`**: Detailed record of employment changes per subject (multiple rows per ID).
- **`Sheet1`**: Main subject list with computed summaries.
- **`Sheet2`**: Filtered list of terminated subjects.
- **`Sheet3`**: IDs of subjects who began as GRN/GRM.

> ⚠️ Terminated subjects were all terminated after the start of the COVID-19 pandemic.

---

## 📊 Variable Definitions

| Variable                         | Description                                             |
|----------------------------------|---------------------------------------------------------|
| `fte change`, `division change`, `award change`, `salary change` | 1 = at least one change, 0 = no change                |
| `start by GRN`                  | 1 = started as GRN/GRM, 0 = otherwise                   |
| `graduate N/M and terminated`   | 1 = was GRN/GRM and terminated, 0 = otherwise           |

---

## ⚙️ How to Use

1. Open your Excel `.xlsm` workbook.
2. Press `Alt + F11` to launch the **VBA Editor**.
3. Insert a new module and paste the macro code.
4. Run macros in order depending on your goal.
---
📬 Contact
Author: Mina Jahandideh
Email: mn.jahandideh@gmail.com
GitHub: @Mina-Jahandideh

📄 License
This code is shared under the MIT License. Please give appropriate credit if used in academic or applied work.
