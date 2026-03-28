from flask import Flask, request, render_template, send_file
import pandas as pd
import os
import re

app = Flask(__name__)

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)


def clean_column_name(col_name):
    col_name = str(col_name).strip().lower()
    col_name = re.sub(r"[^a-zA-Z0-9]+", "_", col_name)
    col_name = re.sub(r"_+", "_", col_name)
    col_name = col_name.strip("_")
    return col_name


def generate_unique_employee_ids(df, id_column="employee_id"):
    if id_column not in df.columns:
        df[id_column] = [f"EMP{str(i + 1).zfill(4)}" for i in range(len(df))]
        return df

    df[id_column] = df[id_column].astype(str).str.strip()
    df[id_column] = df[id_column].replace(["", "nan", "None", "null", "<NA>"], pd.NA)

    used_ids = set()
    final_ids = []
    counter = 1

    for value in df[id_column]:
        if pd.isna(value) or value in used_ids:
            while True:
                new_id = f"EMP{str(counter).zfill(4)}"
                counter += 1
                if new_id not in used_ids:
                    used_ids.add(new_id)
                    final_ids.append(new_id)
                    break
        else:
            used_ids.add(value)
            final_ids.append(value)

    df[id_column] = final_ids
    return df


def format_mixed_date(value):
    if pd.isna(value):
        return pd.NA

    value = str(value).strip()
    if value == "":
        return pd.NA

    # First try normal parsing
    parsed = pd.to_datetime(value, errors="coerce")

    # If failed, try again with dayfirst=True
    if pd.isna(parsed):
        parsed = pd.to_datetime(value, errors="coerce", dayfirst=True)

    if pd.isna(parsed):
        return pd.NA

    return parsed.strftime("%d/%m/%Y")


def clean_excel_data(df):
    # Remove fully empty rows and columns
    df = df.dropna(axis=0, how="all")
    df = df.dropna(axis=1, how="all")

    # Clean column names
    df.columns = [clean_column_name(col) for col in df.columns]

    # Trim spaces from text cells
    for col in df.columns:
        if df[col].dtype == "object":
            df[col] = df[col].astype(str).str.strip()

    # Replace common empty text values with NaN
    df = df.replace(["", "nan", "None", "null"], pd.NA)

    # Remove duplicate rows
    df = df.drop_duplicates()

    # Convert all date-like columns into DD/MM/YYYY
    for col in df.columns:
        if "date" in col or "dob" in col or "date_of_birth" in col:
            df[col] = df[col].apply(format_mixed_date)

    # Salary column: fill missing values with average salary
    if "salary" in df.columns:
        df["salary"] = pd.to_numeric(df["salary"], errors="coerce")
        avg_salary = df["salary"].mean()
        if pd.notna(avg_salary):
            df["salary"] = df["salary"].fillna(round(avg_salary, 2))

    # Age column: fill missing values with average age
    if "age" in df.columns:
        df["age"] = pd.to_numeric(df["age"], errors="coerce")
        avg_age = df["age"].mean()
        if pd.notna(avg_age):
            df["age"] = df["age"].fillna(round(avg_age)).astype(int)

    # Generate unique employee ids
    df = generate_unique_employee_ids(df, "employee_id")

    # Fill remaining missing values
    for col in df.columns:
        if col in ["salary", "age"]:
            continue
        if df[col].dtype == "object":
            df[col] = df[col].fillna("N/A")
        else:
            df[col] = df[col].fillna(0)

    return df


@app.route("/")
def home():
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    if "file" not in request.files:
        return "No file part found"

    file = request.files["file"]

    if file.filename == "":
        return "No file selected"

    if not file.filename.endswith((".xlsx", ".xls")):
        return "Please upload only Excel files (.xlsx or .xls)"

    input_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(input_path)

    try:
        df = pd.read_excel(input_path)
        cleaned_df = clean_excel_data(df)

        output_path = os.path.join(OUTPUT_FOLDER, "cleaned_" + file.filename)
        cleaned_df.to_excel(output_path, index=False)

        return send_file(output_path, as_attachment=True)

    except Exception as e:
        return f"Error while processing file: {str(e)}"


if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)