# Excel Format Requirement
# - Each curve must occupy two columns:
#- Column 1: Displacement (x)
#- Column 2: Nodal Shear (y)
#- For example:
#| X1 | Y1 | X2 | Y2 | X3 | Y3 | ...
#|----|----|----|----|----|----|
#| 0  | 0  | 0  | 0  | 0  | 0  |
#| 5  | 1200 | 5 | 1100 | 5 | 1000 |
#| ... | ... | ... | ... | ... | ... |
print("Is your excel file formatted correctly? 2 column for each curver consequetevly in form of X1|Y1|X2|Y2.... from top to bottom floor (yes/no)")
if input().strip().lower() != "yes":
    print("Please ensure your Excel file follows the specified format.")
    exit()
import openpyxl
import tkinter as tk
from tkinter import filedialog
import matplotlib.pyplot as plt
import numpy as np
from scipy.signal import savgol_filter
from collections import defaultdict
from matplotlib.ticker import MultipleLocator, FuncFormatter
import itertools

# Optional: import your custom fitting logic
try:
    import si  # si.py must be in the same directory
    use_si = True
except ImportError:
    use_si = False

def polynomial_fit(x, y, degree=3):
    x = np.array(x)
    y = np.array(y)
    coeffs = np.polyfit(x, y, degree)
    return np.polyval(coeffs, x)

def get_excel_data(file_path, n_sets):
    wb = openpyxl.load_workbook(file_path)
    sheet = wb.active
    all_x, all_y, labels = [], [], []
    color_cycle = itertools.cycle(plt.rcParams['axes.prop_cycle'].by_key()['color'])

    for i in range(n_sets):
        x_col = 2 * i + 1
        y_col = 2 * i + 2
        x_vals, y_vals = [], []

        for row in sheet.iter_rows(min_row=2):
            x_cell = row[x_col - 1].value
            y_cell = row[y_col - 1].value
            if x_cell is not None and y_cell is not None:
                x_vals.append(float(x_cell))
                y_vals.append(float(y_cell) / 1000)  # Convert to kN

        label = input(f"Enter label for Curve {i+1}:\n")
        color = input(f"Enter color for Curve {i+1}:\n")

        all_x.append(x_vals)
        all_y.append(y_vals)
        labels.append(label)
        color.append(color)

    return all_x, all_y, labels

def plot_initial_curves(all_x, all_y, labels, title):
    plt.figure()
    for x, y, (label, color) in zip(all_x, all_y, labels):
        plt.plot(x, y, linewidth=1, color=color, label=label)

    ax = plt.gca()
    ax.set_xlabel("Displacement (mm)")
    ax.set_ylabel("Nodal Force (kN)")
    ax.set_title(title)
    ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.3)
    ax.legend(loc="upper left")

    flat_y = [val for sub in all_y for val in sub]
    x_max = max([max(x) for x in all_x])
    y_max = max(flat_y)
    ax.set_xlim(0, np.ceil(x_max / 5) * 5)
    ax.set_ylim(0, np.ceil(y_max / 1) * 1 + 10)
    ax.xaxis.set_major_locator(MultipleLocator(5))
    ax.yaxis.set_major_locator(MultipleLocator(5))
    plt.tight_layout()
    plt.show()
model_height=2729.5  # Example model height in mm 
self_weight= 112.82  # Example self weight in kN
def generate_pushover_curve(x1, y1, x2, y2, title):
    if len(x1) != len(x2):
        raise ValueError("Input curves must have the same number of points.")

    x_combined = x1 if max(x1) >= max(x2) else x2
    base_shear = [y1[i] + y2[i] for i in range(len(x_combined))]

    shear_map = defaultdict(list)
    for xv, yv in zip(x_combined, base_shear):
        shear_map[xv].append(yv)

    x_unique = np.array(sorted(shear_map.keys()))
    y_unique = np.array([np.mean(shear_map[x]) for x in x_unique])

    method = input("Choose fitting method: 'savgol', 'regression', or 'si'\n").strip().lower()
    if method == "regression":
        degree = int(input("Enter polynomial degree (e.g., 2, 3, 4):\n"))
        smooth_y = polynomial_fit(x_unique, y_unique, degree)
    else:
        window_size = int(input("Enter Savitzky-Golay window size (odd number):\n"))
        poly_order = int(input("Enter polynomial order:\n"))
        if window_size >= len(y_unique):
            window_size = len(y_unique) - 1 if len(y_unique) % 2 == 0 else len(y_unique)
        if window_size % 2 == 0:
            window_size += 1
        smooth_y = savgol_filter(y_unique, window_length=window_size, polyorder=poly_order)

    compare = input("Compare raw and fitted curves? (yes/no)\n").strip().lower()
    plt.figure()
    ax = plt.gca()
    if compare == "yes":
        ax.plot(x_unique, y_unique, color="gray", linestyle="--", linewidth=1, label="raw pushover")

    ax.plot(x_unique, smooth_y, color="blue", linewidth=1.5, label="smoothen pushover")
    ax.set_xlabel("Top Displacement (mm)")
    ax.set_ylabel("Base Shear (kN)")
    ax.set_title(f"{title} — Pushover Curve")
    ax.grid(True, linestyle="--", linewidth=0.5, alpha=0.3)
    ax.legend(loc="lower right")

    secax = ax.secondary_xaxis('top', functions=(lambda x: (x / model_height) * 100, lambda drift: (drift / 100) * model_height))
    secax.set_xlabel("Drift (%)")
    tick_spacing_mm = 5
    tick_spacing_drift = (tick_spacing_mm / model_height) * 100
    secax.xaxis.set_major_locator(MultipleLocator(tick_spacing_drift))
    secax.xaxis.set_major_formatter(FuncFormatter(lambda val, pos: f"{val:.2f}" if tick_spacing_drift < 5 else f"{val:.1f}"))
    ax.set_xlim(0, np.ceil(x_unique.max() / tick_spacing_mm) * tick_spacing_mm)

    secay = ax.secondary_yaxis('right', functions=(lambda y: (y / self_weight), lambda c: c * self_weight))
    secay.set_ylabel("Base Shear Coefficient")
    tick_spacing_kn = 5
    tick_spacing_bsc = (tick_spacing_kn / self_weight)
    secay.yaxis.set_major_locator(MultipleLocator(tick_spacing_bsc))
    secay.yaxis.set_major_formatter(FuncFormatter(lambda val, pos: f"{val:.2f}" if tick_spacing_bsc < 1 else f"{val:.1f}"))
    ax.set_ylim(0, np.ceil(y_unique.max() / tick_spacing_kn) * tick_spacing_kn)

    plt.tight_layout()
    plt.show()

    export = input("Export plot to file? (yes/no)\n").strip().lower()
    if export == "yes":
        filename = input("Enter filename (without extension):\n")
        fmt = input("Choose format: 'png' or 'pdf'\n").strip().lower()
        plt.savefig(f"{filename}.{fmt}", dpi=300)
        print(f"✅ Plot saved as {filename}.{fmt}")

# Main loop
while True:
    root = tk.Tk()
    root.withdraw()
    file_path = filedialog.askopenfilename(title="Select your Excel file", filetypes=[("Excel files", "*.xlsx")])
    if not file_path:
        print("❌ No file selected.")
        break

    try:
        n_sets = int(input("How many curves are in the Excel file?\n"))
        title = input("Enter title for the initial graph:\n")
        all_x, all_y, labels = get_excel_data(file_path, n_sets)
        plot_initial_curves(all_x, all_y, labels, title)

        if input("Generate pushover curve from first two curves? (yes/no)\n").strip().lower() == "yes":
            generate_pushover_curve(all_x[0], all_y[0], all_x[1], all_y[1], title)

        if input("Replot with new inputs? (yes/no)\n").strip().lower() != "yes":
            print("Exiting program. Goodbye!")
            break

    except Exception as e:
        print(f"⚠️ Error: {e}")
        break