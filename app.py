import streamlit as st
import pdfplumber
import pandas as pd
import re
import io

# --- CONFIGURATION ---
REQUIRED_COLUMNS = [
    "Party Name",
    "Station",
    "Product Name",
    "Qty",
    "Rate",
    "Amount",
    "Tax %"
]

# --- NAME MAPPING ---
PRODUCT_MAPPING = {
    'ALPHALACT LF 200GM': 'AlphaLacT LF 200gm',
    'ALPHALACT PRM-1 400GM': 'AlphaLacT-I 400gm',
    'ALPHALACT PRM-2 400GM': 'AlphaLacT-2 400gm',
    'ALPHAFIT MOM 200GM': 'AlphaFiT MoM 200gm Choc',
    'ALPHALACT PRM-1 200GM': 'AlphaLacT-I 200gm',
    'ALPHALACT LBW 400GM': 'AlphaLacT LBW 400gm',
    'ALPHALACT PLUS-1 400GM': 'AlphaLacT Plus-I 400gm',
    'ALPHALACT PLUS-1 200GM': 'AlphaLacT- Plus-I 200gm'
}


# --- HELPER FUNCTIONS ---
def clean_number_str(s):
    if not s: return ""
    # Remove commas and hyphens attached to numbers (e.g., "12-")
    return s.replace(',', '').replace('-', '').strip()


def parse_number(s):
    try:
        clean = clean_number_str(s).replace('%', '')
        return float(clean)
    except ValueError:
        return 0.0


def is_numeric_item(s):
    if not s: return False
    clean = clean_number_str(s)
    if clean == '' or len(clean) > 15: return False
    # Must allow simple numbers, but reject mixes like '200GM'
    if re.search(r'[a-zA-Z]', clean): return False
    try:
        float(clean)
        return True
    except ValueError:
        return False


def process_pdf(uploaded_file, debug_mode=False):
    all_rows = []
    debug_logs = []

    # Use strict tolerance to keep '200GM' and '2' separate
    EXTRACT_SETTINGS = {
        "x_tolerance": 1,
        "y_tolerance": 5
    }

    with pdfplumber.open(uploaded_file) as pdf:
        for page_num, page in enumerate(pdf.pages, 1):
            words = page.extract_words(**EXTRACT_SETTINGS)

            rows = {}
            for word in words:
                y_bucket = round(word['top'] / 5) * 5
                if y_bucket not in rows: rows[y_bucket] = []
                rows[y_bucket].append(word)

            sorted_y = sorted(rows.keys())

            current_party = ""
            current_station = ""

            for y in sorted_y:
                line_words = sorted(rows[y], key=lambda w: w['x0'])
                line_text_parts = [w['text'] for w in line_words]
                full_text = " ".join(line_text_parts).strip()

                if len(full_text) < 3: continue

                # Filter 'Total' rows
                if "total" in full_text.lower():
                    if debug_mode: debug_logs.append(f"🚫 Skipped 'Total': {full_text}")
                    continue

                if "Page No" in full_text or "VEDIKA PHARMACY" in full_text or "DESCRIPTION" in full_text:
                    continue

                # --- HEADER DETECTION ---
                is_data_row = False
                # Count numbers, ignoring hyphens
                numeric_count = sum(1 for w in line_text_parts if is_numeric_item(w))

                if numeric_count >= 3:
                    is_data_row = True

                if not is_data_row:
                    parts = full_text.split("-")
                    # Be careful not to split "PRM-1" thinking it's a separator
                    # Heuristic: headers usually have spaces around the hyphen or are distinct
                    if len(parts) >= 2 and not any(is_numeric_item(p) for p in parts):
                        current_station = parts[-1].strip()
                        current_party = "-".join(parts[:-1]).strip()
                    else:
                        current_party = full_text
                        current_station = ""
                    if debug_mode: debug_logs.append(f"🏷️ HEADER: {current_party}")
                    continue

                # --- DATA ROW PARSING (IMPROVED LOOP) ---
                numeric_block = []
                name_parts = []

                # Iterate backwards
                for i in range(len(line_text_parts) - 1, -1, -1):
                    word = line_text_parts[i]

                    # SKIP standalone hyphens at the end
                    if word.strip() == '-':
                        continue

                    if is_numeric_item(word):
                        numeric_block.insert(0, parse_number(word))
                    else:
                        # Once we hit text, stop.
                        name_parts = line_text_parts[:i + 1]
                        break

                qty, rate, amount, tax = 0, 0, 0, 0
                count = len(numeric_block)

                if count >= 5:
                    qty = numeric_block[-5]
                    rate = numeric_block[-3]
                    amount = numeric_block[-2]
                    tax = numeric_block[-1]
                elif count == 4:
                    qty = numeric_block[-4]
                    rate = numeric_block[-3]
                    amount = numeric_block[-2]
                    tax = numeric_block[-1]
                elif count == 3:
                    # [Rate?, Amount, Tax] -> Missing Qty
                    amount = numeric_block[-2]
                    tax = numeric_block[-1]
                    rate = numeric_block[0]
                else:
                    if debug_mode: debug_logs.append(f"⚠️ SKIPPED (Low Data): {full_text}")
                    continue

                # --- CLEANUP PRODUCT NAME ---
                raw_product_name = " ".join(name_parts).strip()

                # Regex 1: Clean trailing hyphen
                raw_product_name = raw_product_name.rstrip(' -')

                # Regex 2: Check for trailing number + hyphen pattern (e.g. "Name 12 -")
                # This catches what the loop might have missed if formatting was weird
                match = re.search(r'\s(\d+)\s*-?$', raw_product_name)

                if match:
                    trailing_number = float(match.group(1))

                    # Logic: If Qty is missing/zero, OR if we have a small integer likely to be Qty
                    if qty == 0 or (trailing_number < 1000 and rate > trailing_number):
                        # Fix the Qty
                        qty = trailing_number
                        # Remove it from name
                        raw_product_name = raw_product_name[:match.start()].strip()
                        # Clean trailing hyphen again just in case
                        raw_product_name = raw_product_name.rstrip(' -')
                        if debug_mode: debug_logs.append(f"🔧 Extracted Qty {qty} from name")

                # Apply Mapping
                final_name = PRODUCT_MAPPING.get(raw_product_name, raw_product_name)

                # Fallback: Extract Party from Name
                if not current_party:
                    if " - " in raw_product_name:
                        split_parts = raw_product_name.split(" - ")
                        current_party = split_parts[0]
                        final_name = " - ".join(split_parts[1:])
                        final_name = PRODUCT_MAPPING.get(final_name, final_name)

                all_rows.append({
                    "Party Name": current_party,
                    "Station": current_station,
                    "Product Name": final_name,
                    "Qty": qty,
                    "Rate": rate,
                    "Amount": amount,
                    "Tax %": tax
                })

    return pd.DataFrame(all_rows, columns=REQUIRED_COLUMNS), debug_logs


def to_excel(df):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='openpyxl') as writer:
        df.to_excel(writer, index=False, sheet_name='Sales Data')
    return output.getvalue()


# --- STREAMLIT UI ---
st.set_page_config(page_title="PDF to Excel Converter", page_icon="📊", layout="wide")

st.title("📊 PDF to Excel Converter (Final Cleaned)")
st.markdown("Extracts sales data, fixing 'Name 12 -' issues and ensuring correct Qty.")

col1, col2 = st.columns([2, 1])
with col1:
    uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")
with col2:
    debug_mode = st.checkbox("Show Debug Logs", help="See extracted lines and skipped rows.")

if uploaded_file is not None:
    with st.spinner('Processing...'):
        try:
            df, logs = process_pdf(uploaded_file, debug_mode)

            if debug_mode:
                st.subheader("Debug Logs")
                st.text_area("Log Output", "\n".join(logs), height=200)

            if not df.empty:
                st.success(f"Successfully extracted {len(df)} rows!")
                st.dataframe(df)

                excel_data = to_excel(df)
                st.download_button(
                    label="📥 Download Excel File",
                    data=excel_data,
                    file_name="sales_data_cleaned.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
            else:
                st.error("No data found. Enable 'Debug Logs' to see what went wrong.")

        except Exception as e:
            st.error(f"Error: {e}")