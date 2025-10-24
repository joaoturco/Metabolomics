import pandas as pd
import numpy as np
import itertools
from tqdm import tqdm
import os

# === SETTINGS ===
input_path = r"C:\Users\User\Desktop\Artigo RMN Tese\ADSS\NMR_003.xlsx"  # <<< Put your file here
sheet_name = 0  # index or name of the Excel sheet
output_file = "C:/Users/User/Desktop/Artigo RMN Tese/ADSS/ADSS5.xlsx"

shift_column = 0  # index of the column containing the chemical shift (ppm)

# Define multiple regions of interest in the spectrum (in ppm)
# Example: [(0.5, 1.5), (3.0, 4.5), (6.0, 8.0)]
regions = [(0, 10), (0.5, 2.5), (3.5, 5.5), (6.0, 8.0)]

# === DATA READING ===
df = pd.read_excel(input_path, sheet_name=sheet_name)

ppm = df.iloc[:, shift_column].reset_index(drop=True)  # chemical shifts
data = df.drop(df.columns[shift_column], axis=1).reset_index(drop=True)  # intensities
sample_columns = data.columns


# === NORMALIZATION FUNCTION ===
def normalize_columns(data):
    return data.divide(data.sum(axis=0), axis=1) * 100


# === FUNCTION TO CALCULATE ADSS ===
def calculate_adss(percent_table):
    samples = percent_table.columns
    sim_matrix = pd.DataFrame(index=samples, columns=samples, dtype=float)

    for s1, s2 in tqdm(itertools.product(samples, repeat=2), total=len(samples) ** 2, desc="Calculating ADSS"):
        diff = (percent_table[s1] - percent_table[s2]).abs().sum()
        adss = 100 - diff / 2
        sim_matrix.loc[s1, s2] = round(adss, 2)

    return sim_matrix


# === PROCESS BY REGIONS ===
adss_results = {}

with pd.ExcelWriter(output_file, engine="xlsxwriter") as writer:
    # Save original data
    df.to_excel(writer, sheet_name="Original", index=False)

    # Iterate over defined regions
    for (ppm_min, ppm_max) in regions:
        print(f"\n🔎 Processing region {ppm_min}-{ppm_max} ppm...")

        # Filter spectrum by region
        mask = (ppm >= ppm_min) & (ppm <= ppm_max)
        df_filtered = data.loc[mask].copy()
        ppm_filtered = ppm[mask].reset_index(drop=True)

        # Normalize
        df_norm = normalize_columns(df_filtered).reset_index(drop=True)
        df_norm.insert(0, "ppm", ppm_filtered)

        # Calculate ADSS
        adss_matrix = calculate_adss(df_norm.drop("ppm", axis=1))

        # Region name
        region_name = f"{ppm_min}-{ppm_max}ppm"

        # Save results to Excel
        adss_matrix.to_excel(writer, sheet_name=f"ADSS_{region_name}")
        df_norm.to_excel(writer, sheet_name=f"Norm_{region_name}", index=False)

print(f"\n✅ ADSS analysis completed! Results saved to:\n{output_file}")
