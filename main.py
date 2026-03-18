import streamlit as st
import pandas as pd
import io
import pyexcel as p

def check_password():
    def password_entered():
        if st.session_state["password"] == st.secrets["auth"]["password"]:
            st.session_state["password_correct"] = True
            del st.session_state["password"]
        else:
            st.session_state["password_correct"] = False

    if "password_correct" not in st.session_state:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.stop()

    if not st.session_state["password_correct"]:
        st.text_input("Password", type="password", on_change=password_entered, key="password")
        st.error("Incorrect password")
        st.stop()

check_password()

# -----------------------------
# CONFIG
# -----------------------------
st.set_page_config(page_title="WLI Stock Values", layout="wide")

# -----------------------------
# DEFAULT EXCHANGE RATES
# -----------------------------
default_rates = {
    'USD': 0.85,
    'AUD': 0.60,
    'NZD': 0.51,
    'ZAR': 0.053,
    'GBP': 1.14,
    'EUR': 1
}

# -----------------------------
# FUNCTIONS
# -----------------------------
def safe_convert(row, value_col, currency_col, rates):
    """Safely convert a row's value to EUR."""
    if value_col not in row or currency_col not in row:
        return 0
    currency = str(row[currency_col]).strip()
    value = float(row[value_col]) if pd.notna(row[value_col]) else 0
    rate = rates.get(currency, 1)
    return value * rate


def dedupe_columns(cols):
    """Ensure duplicate column names become Value, Value.1, Value.2, etc."""
    seen = {}
    new_cols = []
    for col in cols:
        if col not in seen:
            seen[col] = 0
            new_cols.append(col)
        else:
            seen[col] += 1
            new_cols.append(f"{col}.{seen[col]}")
    return new_cols


def load_excel_any_format(uploaded_file):
    """Load .xls or .xlsx safely without xlrd."""
    name = uploaded_file.name.lower()

    # Try pyexcel for .xls
    if name.endswith(".xls"):
        try:
            uploaded_file.seek(0)
            sheet = p.get_sheet(file_type="xls", file_content=uploaded_file.read())
            data = sheet.to_array()
            df = pd.DataFrame(data)
            return df
        except Exception:
            pass

    # Try openpyxl for .xlsx
    try:
        uploaded_file.seek(0)
        return pd.read_excel(uploaded_file, header=None)
    except Exception:
        pass

    # Try CSV fallback
    try:
        uploaded_file.seek(0)
        return pd.read_csv(uploaded_file, sep=None, engine="python", header=None)
    except Exception:
        pass

    st.error("This file is not a valid Excel format.")
    return None


# -----------------------------
# UI
# -----------------------------
st.title("WLI Stock Values")

# -----------------------------
# EXCHANGE RATE EDITOR (PERSISTENT)
# -----------------------------
st.subheader("💱 Exchange Rates")

if "exchange_rates" not in st.session_state:
    st.session_state.exchange_rates = default_rates.copy()

rates_df = pd.DataFrame(
    {
        "Currency": list(st.session_state.exchange_rates.keys()),
        "Rate to EUR": list(st.session_state.exchange_rates.values())
    }
)

edited_rates_df = st.data_editor(
    rates_df,
    num_rows="dynamic",
    use_container_width=True
)

st.session_state.exchange_rates = dict(
    zip(edited_rates_df["Currency"], edited_rates_df["Rate to EUR"])
)

exchange_rates = st.session_state.exchange_rates

st.success("Exchange rates updated.")

# -----------------------------
# FILE UPLOADS
# -----------------------------
erp_file = st.file_uploader("ERP Stock File (.xls or .xlsx)", type=["xls", "xlsx"])
insured_file = st.file_uploader("Insured Values File (.xlsx)", type=["xlsx"])

if erp_file and insured_file:
    st.success("Files uploaded successfully.")

    # -----------------------------
    # LOAD ERP FILE
    # -----------------------------
    df = load_excel_any_format(erp_file)
    if df is None:
        st.stop()

    # -----------------------------
    # FORCE ROW 3 AS HEADER
    # -----------------------------
    df.columns = df.iloc[2]
    df = df.iloc[3:].reset_index(drop=True)

    # -----------------------------
    # REMOVE ERP TOTALS BLOCK
    # -----------------------------
    df = df[
        df['Vineyard'].notna() &
        (df['Vineyard'].astype(str).str.strip() != "") &
        (~df['Vineyard'].astype(str).str.strip().str.lower().eq("total")) &
        (~df['Vineyard'].astype(str).str.strip().str.isnumeric())
    ]

    # -----------------------------
    # REMOVE SPECIFIC VINEYARDS
    # -----------------------------
    excluded_vineyards = [
        "Artemis Wines / Von Baron Holdings Pty Ltd",
        "Buzzbox",
        "Jackson Wine Estates International",
        "Overseas Wine Import",
        "The Appletree Cider Company Ltd",
        "Wine Logistics - Destruction 6"
    ]
    df = df[~df['Vineyard'].isin(excluded_vineyards)]

    # -----------------------------
    # CLEAN HEADERS
    # -----------------------------
    df.columns = dedupe_columns(df.columns)
    df.columns = [
        col if isinstance(col, str) and col.strip() != "" else f"Empty_{i}"
        for i, col in enumerate(df.columns)
    ]

    # Remove rows where Vineyard is empty but Currency is filled
    if "Currency" in df.columns:
        df = df[~(
            df['Vineyard'].astype(str).str.strip().eq("") &
            df['Currency'].astype(str).str.strip().ne("")
        )]

    # -----------------------------
    # CONVERT VALUES (robust pairing)
    # -----------------------------
    df['Total_Value_EUR_Row'] = 0

    for i, col in enumerate(df.columns):
        if col.startswith("Value"):
            value_col = col
            currency_col = df.columns[i + 1] if i + 1 < len(df.columns) else None

            df['Total_Value_EUR_Row'] += df.apply(
                lambda row: safe_convert(row, value_col, currency_col, exchange_rates),
                axis=1
            )

    # -----------------------------
    # AGGREGATE
    # -----------------------------
    summary = df.groupby('Vineyard', as_index=False).agg(
        Total_Value_EUR=('Total_Value_EUR_Row', 'sum')
    )

    # -----------------------------
    # LOAD INSURED VALUES
    # -----------------------------
    insured_df = pd.read_excel(insured_file)
    insured_df.columns = insured_df.columns.str.strip()
    insured_df = insured_df[['Vineyard', 'Insured Value']]
    insured_df = insured_df.drop_duplicates(subset='Vineyard', keep='first')

    # -----------------------------
    # MERGE
    # -----------------------------
    final_report = pd.merge(summary, insured_df, on='Vineyard', how='left')

    # -----------------------------
    # ADD INSURED-ONLY VINEYARDS
    # -----------------------------
    extra_vineyards = {
        "Oollin srl": 160120,
        "Mean srl": 80060
    }

    for vineyard, insured_val in extra_vineyards.items():
        if vineyard not in final_report['Vineyard'].values:
            final_report = pd.concat([
                final_report,
                pd.DataFrame({
                    'Vineyard': [vineyard],
                    'Total_Value_EUR': [insured_val],
                    'Insured Value': [insured_val]
                })
            ], ignore_index=True)

    # -----------------------------
    # GRAND TOTAL
    # -----------------------------
    grand_total = pd.DataFrame({
        'Vineyard': ['GRAND TOTAL'],
        'Total_Value_EUR': [final_report['Total_Value_EUR'].sum()],
        'Insured Value': [final_report['Insured Value'].sum()]
    })

    final_report = pd.concat([final_report, grand_total], ignore_index=True)

    # -----------------------------
    # DISPLAY RESULTS
    # -----------------------------
    st.subheader("📊 Final Report")
    st.dataframe(final_report, use_container_width=True)

    # -----------------------------
    # DOWNLOAD BUTTON
    # -----------------------------
    buffer = io.BytesIO()
    final_report.to_excel(buffer, index=False, engine='openpyxl')
    buffer.seek(0)

    st.download_button(
        label="⬇️ Download Excel Report",
        data=buffer,
        file_name="Warehouse_Stock_vs_Insured_EUR.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
