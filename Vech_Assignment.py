# Data Cleaning , Keyword & Issue Category Integration + Visualizations

import sys
print("Python Executable:", sys.executable)

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import os
from langdetect import detect
from deep_translator import GoogleTranslator

# 1. Load dataset
file_path = r"xlsfile"   
df = pd.read_excel(file_path, sheet_name="Sheet1")
print("Original Shape:", df.shape)

# 2. Drop useless columns
drop_cols = [
    "CAMPAIGN_NBR",
    "ENGINE_TRACE_NBR",
    "ENGINE_SOURCE_PLANT",
    "TRANSMISSION_SOURCE_PLANT",
    "COMPLAINT_CD_CSI",
    "NON_CAUSAL_PART_QTY"
]
df = df.drop(columns=[c for c in drop_cols if c in df.columns])

# 3. Handle missing values

# Replace blanks in TRANSMISSION_TRACE_NBR with a number
if "TRANSMISSION_TRACE_NBR" in df.columns:
    replacement_number = 999999
    df["TRANSMISSION_TRACE_NBR"] = df["TRANSMISSION_TRACE_NBR"].fillna(replacement_number)

# Fill numerical with median
num_cols = ["TOTALCOST", "LAST_KNOWN_DELVRY_TYPE_CD", "KM", "REPAIR_AGE", "REPORTING_COST", "LBRCOST"]
for col in num_cols:
    if col in df.columns:
        df[col] = df[col].apply(lambda x: np.nan if isinstance(x, (int, float)) and x < 0 else x)
        df[col] = df[col].fillna(df[col].median())

# Fill categorical with mode
cat_cols = ["CAUSAL_PART_NM", "OPTN_FAMLY_CERTIFICATION", "OPTF_FAMLY_EMISSIOF_SYSTEM", "PLANT", "STATE", "LINE_SERIES"]
for col in cat_cols:
    if col in df.columns and df[col].notna().any():
        df[col] = df[col].fillna(df[col].mode()[0])

# 4. Standardize text columns

text_cols = df.select_dtypes(include=["object"]).columns
for col in text_cols:
    df[col] = df[col].astype(str).str.strip().str.upper()

# Translate only long text columns (average > 20 chars)
long_text_cols = [c for c in text_cols if df[c].dropna().astype(str).str.len().mean() > 20]

def translate_to_english(text):
    try:
        if text and detect(text) != 'en':
            return GoogleTranslator(source='auto', target='en').translate(text)
    except:
        return text
    return text

for c in long_text_cols:
    df[c] = df[c].apply(translate_to_english)


# 5. Keyword Extraction & Categorization
# Define issue mapping
issue_map = {
    "Steering Issue": ["STEERING", "STEERING WHEEL", "STEERING WHEEL REPLACEMENT", "STEERING WHEEL WIRE HARNESS"],
    "Fabric Issue": ["APPLIQUE", "TRIM", "STITCH"],
    "Heating Issue": ["HEATER"],
    "Switch Issue": ["SWITCH"],
    "Electrical Issue": ["WIRING", "MODULE"],
    "Horn Connector Issue": ["HORN CONNECTOR"]
}

# Combine CUSTOMER_VERBATIM + CORRECTION_VERBATIM
df["Combined_Verbatim"] = df["CUSTOMER_VERBATIM"].fillna('') + " " + df["CORRECTION_VERBATIM"].fillna('')

# Function to extract keywords and categories
def extract_keywords_and_categories(text):
    if not isinstance(text, str):
        return [], []
    found_keywords = []
    found_categories = []
    for category, keywords in issue_map.items():
        for kw in keywords:
            if kw in text.upper():
                found_keywords.append(kw)
                found_categories.append(category)
    return list(set(found_keywords)), list(set(found_categories))

# Apply extraction
df[["Keyword_Extracted", "Issue_Category"]] = df["Combined_Verbatim"].apply(
    lambda x: pd.Series(extract_keywords_and_categories(x))
)

# Convert lists to comma-separated strings
df["Keyword_Extracted"] = df["Keyword_Extracted"].apply(lambda x: ", ".join(sorted(x)) if x else "")
df["Issue_Category"] = df["Issue_Category"].apply(lambda x: ", ".join(sorted(x)) if x else "")

# 6. Save to Excel

output_file = "Insightful_Task2.xlsx"
df.to_excel(output_file, index=False)
print(f"Data saved to '{output_file}' with Keyword & Issue Category columns.")


# 7. Visualizations

os.makedirs("plots_task2", exist_ok=True)

# Top Keywords (Top 5)
all_keywords = []
for kws in df["Keyword_Extracted"]:
    if kws:
        all_keywords.extend(kws.split(", "))
keyword_series = pd.Series(all_keywords).value_counts().head(5)

plt.figure(figsize=(8,5))
keyword_series.plot(kind="bar", color="salmon")
plt.xticks(rotation=45, ha="right")
plt.title("Top 5 Keywords")
plt.ylabel("Count")
plt.tight_layout()
plt.savefig("plots_task2/top_keywords.png")
plt.show()

# Top 5 Complaint Codes by Avg Repair Age
if "COMPLAINT_CD" in df.columns and "REPAIR_AGE" in df.columns:
    plt.figure(figsize=(10,6))
    avg_repair_age = df.groupby("COMPLAINT_CD")["REPAIR_AGE"].mean().sort_values(ascending=False).head(5)
    avg_repair_age.plot(kind="bar", color="skyblue")
    plt.title("Top 5 Complaint Codes by Avg Repair Age")
    plt.xlabel("Complaint Code")
    plt.ylabel("Average Repair Age")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig("plots_task2/avg_repair_age_by_complaint.png")
    plt.show()

# Top 5 Repair Age Groups by Avg KM
if "KM" in df.columns and "REPAIR_AGE" in df.columns:
    plt.figure(figsize=(10,6))
    avg_km_by_age = df.groupby("REPAIR_AGE")["KM"].mean().sort_values(ascending=False).head(5)
    avg_km_by_age.plot(kind="bar", color="orange")
    plt.title("Top 5 Repair Age Groups by Avg KM")
    plt.xlabel("Repair Age")
    plt.ylabel("Average KM")
    plt.tight_layout()
    plt.savefig("plots_task2/avg_km_by_repair_age.png")
    plt.show()

# Top 5 Dealers by Total Repair Cost
if "DEALER_NAME" in df.columns and "TOTALCOST" in df.columns:
    plt.figure(figsize=(12,6))
    total_cost_dealer = df.groupby("DEALER_NAME")["TOTALCOST"].sum().sort_values(ascending=False).head(5)
    total_cost_dealer.plot(kind="bar", color="seagreen")
    plt.title("Top 5 Dealers by Total Repair Cost")
    plt.xlabel("Dealer Name")
    plt.ylabel("Total Cost")
    plt.xticks(rotation=45, ha="right")
    plt.tight_layout()
    plt.savefig("plots_task2/total_cost_by_dealer.png")
    plt.show()

# Top 5 Issue Categories
if "Issue_Category" in df.columns:
    plt.figure(figsize=(8,8))
    issue_counts = df["Issue_Category"].value_counts().head(5)
    issue_counts.plot(kind="pie", autopct="%1.1f%%", startangle=140, cmap="tab20")
    plt.title("Top 5 Issue Categories")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig("plots_task2/issue_category_distribution.png")
    plt.show()

# Top 5 Countries by Sales Share
if "COUNTRY_SALE_ISO" in df.columns:
    plt.figure(figsize=(8,8))
    country_counts = df["COUNTRY_SALE_ISO"].value_counts().head(5)
    country_counts.plot(kind="pie", autopct="%1.1f%%", startangle=140, cmap="tab10")
    plt.title("Top 5 Countries by Sales Share")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig("plots_task2/country_sale_iso_distribution.png")
    plt.show()

# Top 5 Dealers by Labor Cost Percentage
if "LBRCOST" in df.columns and "DEALER_NAME" in df.columns:
    plt.figure(figsize=(8,8))
    lbr_cost_by_dealer = df.groupby("DEALER_NAME")["LBRCOST"].sum().sort_values(ascending=False).head(5)
    lbr_cost_by_dealer.plot(kind="pie", autopct="%1.1f%%", startangle=140, cmap="Set3")
    plt.title("Top 5 Dealers by Labor Cost Percentage")
    plt.ylabel("")
    plt.tight_layout()
    plt.savefig("plots_task2/lbrcost_by_dealer_distribution.png")
    plt.show()

print("All plots generated and saved in 'plots_task2' folder.")

