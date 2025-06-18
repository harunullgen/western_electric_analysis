import pandas as pd
import numpy as np


# ------------------- Western Electric Rules -------------------

def rule_1(data, mean, std):
    """
    Rule 1: One point beyond 3σ
    """
    return abs(data[-1] - mean) > 3 * std


def rule_2(data, mean, std):
    """
    Rule 2: Two out of three consecutive points beyond 2σ on the same side
    """
    return (
        abs(data[-1] - mean) > 2 * std and
        abs(data[-2] - mean) > 2 * std and
        np.sign(data[-1] - mean) == np.sign(data[-2] - mean)
    )


def rule_3(data, mean, std):
    """
    Rule 3: Four out of five consecutive points beyond 1σ on the same side
    """
    count = 0
    for i in [-1, -2, -3, -4, -5]:
        if abs(data[i] - mean) > std and np.sign(data[i] - mean) == np.sign(data[-1] - mean):
            count += 1
    return count >= 4


def rule_4(data, mean):
    """
    Rule 4: Eight or more consecutive points on the same side of the mean
    """
    return all(data[-i] > mean for i in range(1, 9)) or all(data[-i] < mean for i in range(1, 9))


def rule_5(data):
    """
    Rule 5: Six points in a row steadily increasing or decreasing
    """
    diffs = np.diff(data[-6:])
    return all(d > 0 for d in diffs) or all(d < 0 for d in diffs)


def rule_6(data):
    """
    Rule 6: Fourteen points in a row alternating up and down
    """
    if len(data) < 15:
        return False
    pattern = [np.sign(data[i] - data[i + 1]) for i in range(-15, -1)]
    return all(pattern[i] != pattern[i + 1] for i in range(len(pattern) - 1))


# ------------------- CLI Entry Point -------------------

if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Western Electric Rule Checker")
    parser.add_argument("input_file", help="Input Excel file path")
    parser.add_argument("output_file", help="Output Excel file path")
    args = parser.parse_args()

    # Read Excel
    df = pd.read_excel(args.input_file)

    # Column separation
    id_cols = ['Store Code', 'Store Name']
    week_cols = [col for col in df.columns if col not in id_cols and "General" not in col]

    # Apply rules
    results = []

    for idx, row in df.iterrows():
        store_id = row['Store Code']
        store_name = row['Store Name']

        series = pd.to_numeric(row[week_cols], errors='coerce').dropna().values

        if len(series) >= 15:
            mean = np.mean(series)
            std = np.std(series)

            r1 = rule_1(series, mean, std)
            r2 = rule_2(series, mean, std)
            r3 = rule_3(series, mean, std)
            r4 = rule_4(series, mean)
            r5 = rule_5(series)
            r6 = rule_6(series)

            results.append({
                "Store Code": store_id,
                "Store Name": store_name,
                "Lower Limit": mean - 3 * std,
                "Upper Limit": mean + 3 * std,
                "Rule 1": r1,
                "Rule 2": r2,
                "Rule 3": r3,
                "Rule 4": r4,
                "Rule 5": r5,
                "Rule 6": r6
            })

    # Save result
    result_df = pd.DataFrame(results)
    result_df.to_excel(args.output_file, index=False)
    print("Western Electric analysis completed.")
