📌 What is this project and why is it important?

In manufacturing and operational environments, detecting early signs of process instability is critical to ensure quality and efficiency. This project leverages Western Electric Rules, a well-established statistical process control methodology, to identify unusual patterns in weekly performance data. It is particularly useful in quality assurance, process monitoring, and operational risk detection for retail chains, production lines, or any environment involving time series data.


# 📊 Western Electric Rules Analyzer

This Python application analyzes weekly store performance data using **Western Electric Rules**, a set of statistical process control (SPC) rules used to detect out-of-control conditions in time series data.

---

## 🧩 What It Does

- Reads an Excel file with weekly metrics per store
- Applies six Western Electric Rules
- Flags any statistical control violations
- Outputs both the original data and the results into a new Excel file

---

## 📁 Input Format

The input Excel file should have the following structure:

| Store Code | Store Name | Week 1 | Week 2 | ... | Week 52 |
|------------|-------------|--------|--------|-----|---------|
| 1001       | Store A     | 1.02   | 0.98   | ... | -0.12   |
| 1002       | Store B     | -0.55  | -0.45  | ... | -0.39   |

- Columns must include `Store Code` and `Store Name`
- Weekly columns should be named as `"Week 1"`, `"Week 2"`, ..., `"Week 52"`

---

## 🧪 Western Electric Rules Implemented

| Rule      | Description                                                                  |
|-----------|------------------------------------------------------------------------------|
| **Rule 1** | One point beyond ±3σ                                                        |
| **Rule 2** | Two of three consecutive points beyond ±2σ on the same side                 |
| **Rule 3** | Four of five points beyond ±1σ on the same side                             |
| **Rule 4** | Eight consecutive points on the same side of the mean                       |
| **Rule 5** | Six points steadily increasing or decreasing                                |
| **Rule 6** | Fourteen points alternating up and down                                     |

---

## ⚙️ How to Run

To execute the program, follow these steps in your terminal:

```bash
# 1. Install dependencies
pip install -r requirements.txt

# 2. Run the script
python main.py sample_input.xlsx result_output.xlsx

# 3. Get help (optional)
python main.py --help
