import pandas as pd

# Input Data
data = {
    "Alternatif": ["A1", "A2", "A3", "A4", "BOBOT"],
    "C1 (Juta Rp) Cost": [150, 500, 200, 350, 5],
    "C2 (%) Benefit": [15, 200, 10, 100, 3],
    "C3 Benefit": [2, 2, 3, 3, 2],
    "C4 Benefit": [2, 3, 1, 1, 5],
    "C15 Benefit": [3, 2, 3, 2, 4]
}

# Create a DataFrame for Input Data
df_input = pd.DataFrame(data)

# Weight Normalization
weights = {
    "Criteria": ["C1", "C2", "C3", "C4", "C15"],
    "Weight": [5, 3, 2, 5, 4],
    "Total Weight": [19, 19, 19, 19, 19],
    "Normalized Weight": [0.263, 0.158, 0.105, 0.263, 0.211],
    "Type": ["Cost (-)", "Benefit (+)", "Benefit (+)", "Benefit (+)", "Benefit (+)"]
}

# Create a DataFrame for Weight Normalization
df_weights = pd.DataFrame(weights)

# Vector S Calculation
vector_s = {
    "Alternative": ["A1", "A2", "A3", "A4"],
    "C1^(-w1)": [0.299, 0.207, 0.272, 0.224],
    "C2^(w2)": [1.453, 2.414, 1.320, 1.978],
    "C3^(w3)": [1.074, 1.074, 1.116, 1.116],
    "C4^(w4)": [1.278, 1.346, 1.000, 1.000],
    "C15^(w5)": [1.264, 1.154, 1.264, 1.154],
    "Vector S": [0.691, 0.954, 0.507, 0.575]
}

# Create a DataFrame for Vector S
df_vector_s = pd.DataFrame(vector_s)

# Vector V Calculation
vector_v = {
    "Alternative": ["A1", "A2", "A3", "A4"],
    "Vector S": [0.691, 0.954, 0.507, 0.575],
    "Sum of Vector S": [2.727] * 4,
    "Vector V (S/Sum)": [0.253, 0.350, 0.186, 0.211],
    "Rank": [2, 1, 4, 3]
}

# Create a DataFrame for Vector V
df_vector_v = pd.DataFrame(vector_v)

# Final Result
final_result = {
    "Rank": [1, 2, 3, 4],
    "Alternative": ["A2", "A1", "A4", "A3"],
    "Vector V Value": [0.350, 0.253, 0.211, 0.186]
}

# Create a DataFrame for Final Result
df_final_result = pd.DataFrame(final_result)

# Create an Excel writer
with pd.ExcelWriter("Weighted_Product_Method.xlsx", engine="openpyxl") as writer:
    # Write DataFrames to different sheets
    df_input.to_excel(writer, sheet_name="Input Data", index=False)
    df_weights.to_excel(writer, sheet_name="Weight Normalization", index=False)
    df_vector_s.to_excel(writer, sheet_name="Vector S", index=False)
    df_vector_v.to_excel(writer, sheet_name="Vector V", index=False)
    df_final_result.to_excel(writer, sheet_name="Final Result", index=False)

print("Excel file has been created successfully!")
