import streamlit as st
import pdfplumber
import pandas as pd
import io
import os
import re
from st_aggrid import AgGrid, GridOptionsBuilder
import plotly.express as px

st.set_page_config(page_title="Document Extractor", page_icon="📊")
st.image("https://www.answerfinancial.com/ContentResponsive/Assets/images/partners/partners-page/plymouthrock.png", width=300)
st.title("Product Management - Document Extractor 📊")

# Utils
def dedup_columns(cols):
    seen = {}
    new_cols = []
    for col in cols:
        col_str = str(col)
        if col_str in seen:
            seen[col_str] += 1
            new_cols.append(f"{col_str}_{seen[col_str]}")
        else:
            seen[col_str] = 0
            new_cols.append(col_str)
    return new_cols

def name_and_order_sheets(dfs):
    seen = {}
    named = []
    auto = []
    for i, df in enumerate(dfs):
        name = str(df.columns[0]) if df.columns.size > 0 else f"Sheet_{i+1}"
        name = re.sub(r'[^A-Za-z0-9_ ]+', '', name).strip().replace(" ", "_")[:30]
        if not name or name.lower().startswith("table"):
            name = f"Table_{i+1}"
            auto.append((name, df))
        else:
            if name in seen:
                seen[name] += 1
                name = f"{name}_{seen[name]}"
            else:
                seen[name] = 1
            named.append((name, df))
    return named + auto

def save_excel(sheets):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        used_names = set()
        for name, df in sheets:
            original = name
            count = 1
            while name in used_names:
                name = f"{original}_{count}"
                count += 1
            used_names.add(name)
            df.to_excel(writer, sheet_name=name[:31], index=False)
    output.seek(0)
    return output

def extract_tables(pdf_file, page_range=None):
    dataframes = []
    with pdfplumber.open(pdf_file) as pdf:
        pages = pdf.pages
        if page_range:
            try:
                start, end = map(int, page_range.split('-'))
                pages = pages[start-1:end]
            except:
                pass
        progress = st.progress(0)
        for i, page in enumerate(pages):
            table = page.extract_table()
            if table:
                df = pd.DataFrame(table)
                df = df.reset_index(drop=True)
                if df.shape[0] > 2:
                    header_row = df.iloc[0].tolist()
                    next_row = df.iloc[1].tolist()
                    new_header = [next_row[i] if (val is None or str(val).lower() == 'none') else val for i, val in enumerate(header_row)]
                    df.columns = dedup_columns(new_header)
                    df = df.iloc[2:].reset_index(drop=True)
                else:
                    df.columns = dedup_columns(df.columns)
                dataframes.append(df)
            progress.progress((i + 1) / len(pages))
    return dataframes

# --- UI ---
pdf_file = st.file_uploader("Upload a PDF", type=["pdf"])

if pdf_file is not None:
    st.success("✅ File uploaded successfully!")
    page_range = st.text_input("Select page range (e.g., 1-10)")

    if st.button("📤 Extract Tables"):
        with st.spinner("Extracting tables..."):
            tables = extract_tables(pdf_file, page_range)
            sheets = name_and_order_sheets(tables)
            st.session_state["sheets"] = sheets
            st.session_state["plot_ready"] = False
            st.session_state["excel"] = save_excel(sheets)

    if "excel" in st.session_state:
        pdf_base = os.path.splitext(pdf_file.name)[0]
        file_range = f"_pages_{page_range}" if page_range else ""
        file_name = f"{pdf_base}{file_range}.xlsx"
        st.download_button(
            label="📥 Download Excel File",
            data=st.session_state["excel"],
            file_name=file_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        if st.button("📊 Plot Data"):
            st.session_state["plot_ready"] = True

if st.session_state.get("plot_ready") and "sheets" in st.session_state:
    st.subheader("📈 Plot Builder & Sheet Editor")
    sheet_names = list(dict.fromkeys([name for name, _ in st.session_state["sheets"]]))
    show_all = st.checkbox("Show all sheets", value=False)
    options = sheet_names if show_all else [s for s in sheet_names if not re.match(r"Table_\\d+", s)]

    selected_merge = st.multiselect("Merge sheets for plotting", options=options)
    if st.button("🔀 Merge Selected Sheets") and selected_merge:
        merged_df = pd.concat([dict(st.session_state["sheets"])[s] for s in selected_merge], ignore_index=True)
        st.session_state["plot_sheet"] = ("MergedSheet", merged_df)

    if "plot_sheet" not in st.session_state:
        selected = st.selectbox("Select sheet to edit and plot", options=options)
        df = dict(st.session_state["sheets"])[selected].copy()
        sheet_name = selected
    else:
        sheet_name, df = st.session_state["plot_sheet"]

    df.columns = dedup_columns(df.columns)
    st.markdown("**📝 Editable Table View:**")
    gb = GridOptionsBuilder.from_dataframe(df)
    gb.configure_default_column(editable=True)
    gb.configure_selection("multiple", use_checkbox=True)
    grid_options = gb.build()
    grid_response = AgGrid(df, gridOptions=grid_options, update_mode='MODEL_CHANGED')
    df_edited = grid_response["data"]
    selected_rows = grid_response["selected_rows"]

    if st.button("🗑 Delete Selected Rows") and len(selected_rows) > 0:
        df_edited = df_edited.drop(index=[r['index'] for r in selected_rows]).reset_index(drop=True)

    st.markdown("### 📊 Plot Setup")
    x_col = st.multiselect("Select X-axis", df_edited.columns)
    y_col = st.multiselect("Select Y-axis", df_edited.columns)
    chart_type = st.selectbox("Chart Type", ["line", "bar", "scatter", "box"])

    if st.button("📈 Show Plot"):
        if x_col and y_col:
            try:
                fig = px.line(df_edited, x=x_col[0], y=y_col) if chart_type == "line" else \
                      px.bar(df_edited, x=x_col[0], y=y_col) if chart_type == "bar" else \
                      px.scatter(df_edited, x=x_col[0], y=y_col) if chart_type == "scatter" else \
                      px.box(df_edited, x=x_col[0], y=y_col)
                st.plotly_chart(fig)
            except Exception as e:
                st.error(f"Plot failed: {e}")

    edited_file_name = f"{pdf_base}_edited_{sheet_name}.csv"
    st.download_button(
        "💾 Download Edited Sheet",
        data=df_edited.to_csv(index=False).encode("utf-8"),
        file_name=edited_file_name,
        mime="text/csv"
    )
