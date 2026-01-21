import pandas as pd
import json

# Load Excel
df = pd.read_excel("project_emails_with_content.xlsx")

# Convert all columns to Python native types
df = df.astype(str)  # forces all values to string (safe for JSON)

# Convert DataFrame to list of dicts
data = df.to_dict(orient='records')

# Serialize to JSON
json_data = json.dumps(data)  # no int64 issues now

# Optionally, save to JSON file
with open("emails_data.json", "w", encoding="utf-8") as f:
    f.write(json_data)
