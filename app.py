import streamlit as st
import json
import base64
import re
import io
import pandas as pd
from io import BytesIO
from openai import OpenAI
from PIL import Image, ImageEnhance
import plotly.express as px

st.set_page_config(page_title="Bulk Bill Genie Pro", layout="wide", page_icon="📦")

st.markdown("""
<style>
html, body { font-family: Arial, sans-serif; }
.stButton>button {
    width: 100%;
    border-radius: 8px;
    height: 3.5em;
    background-color: #007BFF;
    color: white;
    font-weight: bold;
    font-size: 1rem;
    border: none;
}
.metric-card {
    background-color: #f0f4ff;
    border: 1px solid #c0d0f0;
    border-radius: 12px;
    padding: 20px;
    text-align: center;
}
.metric-card h2 { color: #0056D2; margin: 0; font-size: 2rem; }
.metric-card p { color: #555; margin: 4px 0 0; font-size: 0.85rem; }
</style>
""", unsafe_allow_html=True)

if "raw_df" not in st.session_state:
    st.session_state.raw_df = None
if "bill_meta" not in st.session_state:
    st.session_state.bill_meta = []

SYSTEM_PROMPT = """
You are an expert Accountant AI. Extract all line items from bills including HANDWRITTEN ones.
Rules:
- For handwritten bills watch for ambiguous digits like 1/7 and 0/6 and 3/8. Choose the value that makes math sense.
- Return ONLY valid JSON with no markdown and no explanation.
Output format:
{
  "bill_number": "...",
  "date": "...",
  "vendor": "...",
  "items": [
    {"item_name": "...", "qty": 0, "unit_price": 0.0, "amount": 0.0}
  ],
  "total_amount": 0.0,
  "tax": 0.0,
  "is_handwritten": false,
  "confidence": "high"
}
If any field is missing use empty string for text and 0 for numbers.
Convert all currency symbols and commas before returning numbers.
"""

def prepare_image(uploaded_file, contrast, sharpness):
    img = Image.open(uploaded_file)
    if img.mode in ("RGBA", "P", "LA"):
        bg = Image.new("RGB", img.size, (255, 255, 255))
        if img.mode == "P":
            img = img.convert("RGBA")
        bg.paste(img, mask=img.split()[-1] if img.mode in ("RGBA", "LA") else None)
        img = bg
    elif img.mode != "RGB":
        img = img.convert("RGB")
    max_dim = 1600
    if max(img.size) > max_dim:
        img.thumbnail((max_dim, max_dim), Image.LANCZOS)
    img = ImageEnhance.Contrast(img).enhance(contrast)
    img = ImageEnhance.Sharpness(img).enhance(sharpness)
    buf = io.BytesIO()
    img.save(buf, format="JPEG", quality=92)
    return base64.b64encode(buf.getvalue()).decode("utf-8")

def clean_amount(value):
    if isinstance(value, (int, float)):
        return float(value)
    if isinstance(value, str):
        cleaned = re.sub(r"[^\d.]", "", value)
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    return 0.0

def normalize_items(items):
    col_map = {
        "item_name": "item", "description": "item", "product": "item", "name": "item",
        "qty": "quantity", "count": "quantity", "units": "quantity",
        "price": "unit_price", "rate": "unit_price", "cost": "unit_price",
        "total": "amount", "total_amount": "amount", "line_total": "amount",
    }
    result = []
    for row in items:
        new_row = {}
        for k, v in row.items():
            mapped = col_map.get(k.lower().strip(), k.lower().strip())
            new_row[mapped] = v
        result.append(new_row)
    return result

def parse_response(content):
    content = re.sub(r"```(?:json)?", "", content).strip().rstrip("`").strip()
    data = json.loads(content)
    if isinstance(data, list):
        return data, {}
    if isinstance(data, dict):
        meta = {k: v for k, v in data.items() if k != "items"}
        for key in ["items", "data", "results", "line_items"]:
            if key in data and isinstance(data[key], list):
                return data[key], meta
        return [data], {}
    raise ValueError("Unknown JSON structure")

def verify_math(df):
    if df.empty:
        return df
    for col in ["quantity", "unit_price", "amount"]:
        if col in df.columns:
            df[col] = df[col].apply(clean_amount)
        else:
            df[col] = 0.0
    df["Expected"] = df["quantity"] * df["unit_price"]
    df["Check"] = df.apply(
        lambda x: "OK" if abs(x["Expected"] - x["amount"]) < 0.5 else "MISMATCH", axis=1
    )
    return df

def build_excel(df, meta_list):
    output = BytesIO()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        df.to_excel(writer, sheet_name="All Items", index=False)
        if "Source_File" in df.columns and "amount" in df.columns:
            summary = df.groupby("Source_File").agg(
                Total_Items=("item", "count"),
                Total_Amount=("amount", "sum")
            ).reset_index()
            if meta_list:
                mdf = pd.DataFrame(meta_list)
                summary = summary.merge(mdf, on="Source_File", how="left")
            summary.to_excel(writer, sheet_name="Summary", index=False)
        if "Check" in df.columns:
            mis = df[df["Check"] == "MISMATCH"]
            if not mis.empty:
                mis.to_excel(writer, sheet_name="Mismatches", index=False)
        wb = writer.book
        hfmt = wb.add_format({"bold": True, "bg_color": "#007BFF", "font_color": "white"})
        for sname in writer.sheets:
            ws = writer.sheets[sname]
            ws.set_column(0, 15, 20)
    output.seek(0)
    return output.getvalue()

with st.sidebar:
    st.markdown("## Settings")
    api_key = st.secrets.get("AI_API_KEY", "")
    if not api_key:
        st.error("API key set nahi hai. App owner ko bataiye.")
    else:
        st.success("App ready hai - aap seedha bills upload kar sakte hain.")
    st.markdown("---")
    st.markdown("### Image Enhancement")
    st.caption("Handwritten bills ke liye values badhayein")
    contrast = st.slider("Contrast", 1.0, 3.0, 1.8, 0.1)
    sharpness = st.slider("Sharpness", 1.0, 3.0, 1.5, 0.1)
    st.markdown("---")
    st.markdown("### How to Use")
    st.markdown("1. API Key enter karein")
    st.markdown("2. Bills upload karein")
    st.markdown("3. Process All Bills dabayein")
    st.markdown("4. Dashboard dekhein")
    st.markdown("5. Excel download karein")
    st.markdown("---")
    if st.session_state.raw_df is not None:
        if st.button("Clear All Data"):
            st.session_state.raw_df = None
            st.session_state.bill_meta = []
            st.rerun()

st.markdown("# Bulk Bill Genie Pro")
st.markdown("AI powered bill extraction - printed aur handwritten dono support karta hai.")
st.markdown("---")

uploaded_files = st.file_uploader(
    "Bills upload karein (JPG, PNG)",
    type=["jpg", "png", "jpeg"],
    accept_multiple_files=True
)

if uploaded_files:
    st.markdown("### Uploaded Bills Preview")
    cols = st.columns(min(len(uploaded_files), 4))
    for i, f in enumerate(uploaded_files):
        with cols[i % 4]:
            st.image(f, caption=f.name, use_container_width=True)

    col1, col2 = st.columns([3, 1])
    with col1:
        process_btn = st.button("Process All Bills")
    with col2:
        st.metric("Files Selected", len(uploaded_files))

    if process_btn:
        if not api_key:
            st.error("Pehle sidebar mein API Key enter karein!")
        else:
            client = OpenAI(api_key=api_key, base_url="https://api.aicredits.in/v1")
            all_results = []
            all_meta = []
            errors = []
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.info("Processing: " + file.name + " (" + str(i+1) + "/" + str(len(uploaded_files)) + ")")
                try:
                    b64 = prepare_image(file, contrast, sharpness)
                    response = client.chat.completions.create(
                        model="gpt-4o",
                        messages=[
                            {"role": "system", "content": SYSTEM_PROMPT},
                            {"role": "user", "content": [
                                {"type": "text", "text": "Extract all data from this bill."},
                                {"type": "image_url", "image_url": {"url": "data:image/jpeg;base64," + b64}}
                            ]}
                        ],
                        max_tokens=2000,
                        temperature=0
                    )
                    content = response.choices[0].message.content
                    items, meta = parse_response(content)
                    items = normalize_items(items)
                    for item in items:
                        item["Source_File"] = file.name
                        all_results.append(item)
                    all_meta.append({
                        "Source_File": file.name,
                        "Bill_Number": meta.get("bill_number", ""),
                        "Date": meta.get("date", ""),
                        "Vendor": meta.get("vendor", ""),
                        "Is_Handwritten": meta.get("is_handwritten", False),
                        "Confidence": meta.get("confidence", ""),
                        "Bill_Total": clean_amount(meta.get("total_amount", 0)),
                        "Tax": clean_amount(meta.get("tax", 0)),
                    })
                except json.JSONDecodeError as e:
                    errors.append(file.name + ": JSON error - " + str(e))
                except Exception as e:
                    errors.append(file.name + ": " + str(e))
                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.empty()

            if all_results:
                df = pd.DataFrame(all_results)
                df = verify_math(df)
                st.session_state.raw_df = df
                st.session_state.bill_meta = all_meta
                st.success("Sabhi bills successfully process ho gaye!")

            if errors:
                st.warning("Kuch files mein errors aaye:")
                for err in errors:
                    st.warning(err)

            if not all_results:
                st.error("Koi bhi bill process nahi hua. API key aur images check karein.")

if st.session_state.raw_df is not None:
    df = st.session_state.raw_df
    meta_list = st.session_state.bill_meta

    st.markdown("---")
    st.markdown("### Dashboard")

    total_amount = df["amount"].sum() if "amount" in df.columns else 0
    total_items = len(df)
    total_bills = df["Source_File"].nunique() if "Source_File" in df.columns else 1
    mismatch_count = (df["Check"] == "MISMATCH").sum() if "Check" in df.columns else 0

    c1, c2, c3, c4 = st.columns(4)
    with c1:
        st.markdown('<div class="metric-card"><h2>' + str(total_bills) + '</h2><p>Bills Processed</p></div>', unsafe_allow_html=True)
    with c2:
        st.markdown('<div class="metric-card"><h2>' + str(total_items) + '</h2><p>Total Line Items</p></div>', unsafe_allow_html=True)
    with c3:
        st.markdown('<div class="metric-card"><h2>Rs ' + f"{total_amount:,.0f}" + '</h2><p>Total Amount</p></div>', unsafe_allow_html=True)
    with c4:
        color = "#dc3545" if mismatch_count > 0 else "#28a745"
        st.markdown('<div class="metric-card"><h2 style="color:' + color + '">' + str(mismatch_count) + '</h2><p>Mismatches</p></div>', unsafe_allow_html=True)

    st.markdown("")

    if meta_list:
        st.markdown("### Bill Details")
        st.dataframe(pd.DataFrame(meta_list), use_container_width=True)

    col_c1, col_c2 = st.columns(2)
    with col_c1:
        if "Source_File" in df.columns and "amount" in df.columns:
            bs = df.groupby("Source_File")["amount"].sum().reset_index()
            bs.columns = ["Bill", "Total Amount"]
            fig1 = px.bar(bs, x="Bill", y="Total Amount", title="Amount per Bill",
                          color="Total Amount", color_continuous_scale="Blues")
            fig1.update_layout(plot_bgcolor="rgba(0,0,0,0)", paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig1, use_container_width=True)

    with col_c2:
        if "item" in df.columns and "amount" in df.columns:
            top = df.groupby("item")["amount"].sum().nlargest(8).reset_index()
            fig2 = px.pie(top, names="item", values="amount", title="Top Items by Amount", hole=0.4)
            fig2.update_layout(paper_bgcolor="rgba(0,0,0,0)")
            st.plotly_chart(fig2, use_container_width=True)

    st.markdown("### Extracted Data")

    if "Source_File" in df.columns:
        files = ["All"] + list(df["Source_File"].unique())
        selected = st.selectbox("Bill filter karein", files)
        show_df = df if selected == "All" else df[df["Source_File"] == selected]
    else:
        show_df = df

    only_mis = st.checkbox("Sirf Mismatches dikhao")
    if only_mis and "Check" in show_df.columns:
        show_df = show_df[show_df["Check"] == "MISMATCH"]

    def highlight(row):
        if "Check" in row and row["Check"] == "MISMATCH":
            return ["background-color: #fff3cd"] * len(row)
        return [""] * len(row)

    st.dataframe(show_df.style.apply(highlight, axis=1), use_container_width=True, height=400)

    if mismatch_count > 0:
        st.warning(str(mismatch_count) + " items mein math mismatch hai - inhe manually verify karein.")

    st.markdown("### Export")
    col_d1, col_d2 = st.columns(2)
    with col_d1:
        excel_data = build_excel(df, meta_list)
        st.download_button(
            label="Download Excel Report",
            data=excel_data,
            file_name="Bills_Report.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )
    with col_d2:
        csv_data = df.to_csv(index=False).encode("utf-8")
        st.download_button(
            label="Download CSV",
            data=csv_data,
            file_name="Bills.csv",
            mime="text/csv",
            use_container_width=True
        )

else:
    st.markdown("""
    <div style="text-align:center; padding: 60px 20px; color: #888;">
        <h1>📂</h1>
        <h3>Abhi koi data nahi hai</h3>
        <p>Upar bills upload karein aur Process All Bills dabayein.</p>
    </div>
    """, unsafe_allow_html=True)