import streamlit as st
import pandas as pd
import numpy as np
import time
import re
import io
import zipfile
from io import BytesIO

def format_rupiah(x):
    if pd.isna(x):
        return ""
    # pastikan bisa diubah ke float
    try:
        x = float(x)
    except:
        return x  # biarin apa adanya kalau bukan angka

    # kalau tidak punya desimal (misal 7000.0), tampilkan tanpa ,00
    if x.is_integer():
        formatted = f"{int(x):,}".replace(",", ".")
    else:
        formatted = f"{x:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        # hapus ,00 kalau desimalnya 0 semua (misal 7000,00 → 7000)
        if formatted.endswith(",00"):
            formatted = formatted[:-3]
    return formatted

def highlight_total(row):
    if any(str(x).strip().upper() == "TOTAL" for x in row):
        return ["font-weight: bold; background-color: #D9EAD3; color: #1A5E20;"] * len(row)
    else:
        return [""] * len(row)
    
def highlight_1st_2nd(row, columns):
    styles = [""] * len(columns)
    first_vendor = row.get("1st Vendor")
    second_vendor = row.get("2nd Vendor")

    for i, col in enumerate(columns):
        if col == first_vendor:
            styles[i] = "background-color: #C6EFCE; color: #006100;"
        elif col == second_vendor:
            styles[i] = "background-color: #FFEB9C; color: #9C6500;"
    return styles

st.markdown(
    """
    <div style="font-size:1.75rem; font-weight:700; margin-bottom:9px">
        🧑‍🏫 User Guide: TCO Comparison Round by Round
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    ":red-badge[Indosat] :orange-badge[Ooredoo] :green-badge[Hutchison]"
)
st.caption("INSPIRE 2025 | Oktaviana Sadama Nur Azizah")

# Divider custom
st.markdown(
    """
    <hr style="margin-top:-5px; margin-bottom:10px; border: none; height: 2px; background-color: #ddd;">
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="
        display: flex;
        align-items: center;
        height: 65px;
        margin-bottom: 10px;
    ">
        <div style="text-align: justify; font-size: 15px;">
            <span style="color: #26BDAD; font-weight: 800;">
            TCO Comparison Round by Round</span>
            evaluates TCO changes across negotiation rounds to track vendor pricing progress and
            identify the most competitive movements.
        </div>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("#### Input Structure")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px;">
            The input file required for this menu should be a 
            <span style="color: #FF69B4; font-weight: 500;">multi-file containing single sheet</span>, in eather 
            <span style="background:#C6EFCE; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xlsx</span> or 
            <span style="background:#FFEB9C; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">.xls</span> format. 
            The file name represents the 
            <span style="font-weight: bold;">"ROUND"</span>, 
            and the table structure in each file is as follows:
        </div>
    """,
    unsafe_allow_html=True
)

# Dataframe
columns = ["Scope", "Desc", "Vendor A", "Vendor B", "Vendor C"]
df = pd.DataFrame([[""] * len(columns) for _ in range(3)], columns=columns)

st.dataframe(df, hide_index=True)

# Buat DataFrame 1 row
st.markdown("""
<table style="width: 100%; border-collapse: collapse; table-layout: fixed; font-size: 15px;">
    <tr>
        <td style="border: 1px solid gray; width: 15%;">Sheet1</td>
        <td style="border: 1px solid gray; font-style: italic; color: #26BDAD">single sheet only</td>
    </tr>
</table>
""", unsafe_allow_html=True)

st.markdown("###### Description:")
st.markdown(
    """
    <div style="font-size:15px;">
        <ul>
            <li>
                <span style="display:inline-block; width:100px;">Scope & Desc</span>: non-numeric columns
            </li>
            <li>
                <span style="display:inline-block; width:100px;">Vendor A - C</span>: numeric columns
            </li>
        </ul>
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            The system accommodates a 
            <span style="font-weight: bold;">dynamic table</span>, allowing users to enter any number of non-numeric and numeric columns. 
            Users have the freedom to name the columns as they wish. The system logic relies on 
            <span style="font-weight: bold;">column indices</span>,
            not specific column names.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
            Originally, the table included a 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">TOTAL ROW</span> because this menu is an extension of the 
            <span style="color: #FF69B4; font-weight: 700;">'TCO Comparison by Year → TCO Summary'</span> menu, which analyzed price trends across rounds. However, you are allowed to 
            <span style="color: #ED1C24; font-weight: 700;">EXCLUDE</span> 
            the TOTAL row, since the system will automatically generate its own TOTAL row.
        </div>
    """,
    unsafe_allow_html=True
)

st.divider()
st.markdown("#### Constraint")

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -10px">
            To ensure this menu works correctly, users need to follow certain rules regarding
            the dataset structure.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MULTIPLE FILE NAME]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 15px; margin-top: -10px">
            This menu operates using 
            <span style="color: #FF69B4; font-weight: 500;">multiple files</span>, where each filename is extracted and used as the value for the 
            <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">ROUND</span> column. 
            Therefore, please ensure that each filename correctly represents its corresponding round and 
            <span style="color: #ED1C24; font-weight: bold;">AVOID</span> using ambiguous names.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px;">
            Because the filenames are parsed into a non-numeric column, the system uses 
            <span style="color: #FF69B4; font-weight: 700;">REGEX</span> to detect and sort the rounds in the correct order. 
            For example:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px;">
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R2.xlsx</span>  |
            <span style="background: #FF00AA; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R4.xlsx</span>  |
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R3.xlsx</span>  |
            <span style="background: #FF00AA; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">L2R1.xlsx</span>
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 15px">
            will automatically be sorted as: 
            <span style="font-weight: bold;">L2R1</span> →
            <span style="font-weight: bold;">L2R2</span> →
            <span style="font-weight: bold;">L2R3</span> →
            <span style="font-weight: bold;">L2R4</span>
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 3px; font-weight: bold;">
            Why is this important?
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px;">
            Because the order of the rounds directly affects the analysis of 
            <span style="color: #38CC8A; font-weight: 700;">PRICE MOVEMENT</span>. It is highly recommended to use clear and consistent naming, such as 
            <span style="background:#FF9A09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Round 1</span>
            <span style="background:#FF9A09; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">Round 2</span> 
            and so on.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:orange-badge[2. COLUMN ORDER]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top: -10px">
            When creating tables, it is important to follow the specified column structure. Columns 
            <span style="font-weight: bold;">must</span> be arranged in the following order:
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: center; font-size: 15px; margin-bottom: 10px; font-weight: bold">
            Non-Numeric Columns → Numeric Columns
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px">
            this order is <span style="color: #FF69B4; font-weight: 700;">strict</span> and 
            <span style="color: #FF69B4; font-weight: 700;">cannot be altered</span>!
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:green-badge[3. NUMBER COLUMN]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Please refer the table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["No", "Scope", "Desc", "Vendor A", "Vendor B", "Vendor C"]
data = [
    [1] + [""] * (len(columns) - 1),
    [2] + [""] * (len(columns) - 1),
    [3] + [""] * (len(columns) - 1)
]
df = pd.DataFrame(data, columns=columns)

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not allowed</span> because it contains a 
            <span style="font-weight: bold;">"No"</span> column. 
            The "No" column is prohibited in this menu, as it will be treated as a numeric column by the system, 
            which violates the constraint described in point 2.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:blue-badge[4. FLOATING TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Floating tables are allowed, meaning tables 
            <span style="color: #FF69B4; font-weight: 700;">do not need to start from cell A1</span>. 
            However, ensure that the cells above and to the left of the table are empty, as shown in the example below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["", "A", "B", "C", "D", "E", "F"]

# Buat 5 baris kosong
df = pd.DataFrame([[""] * len(columns) for _ in range(6)], columns=columns)

# Isi kolom pertama dengan 1–6
df.iloc[:, 0] = [1, 2, 3, 4, 5, 6]

# Header bagian kedua
df.loc[1, ["B", "C", "D", "E"]] = ["Desc", "Vendor A", "Vendor B", "Vendor C"]

# Data Software & Hardware
df.loc[2, ["B", "C", "D", "E"]] = ["Optical Cable", "1.000", "2.000", "3.000"]
df.loc[3, ["B", "C", "D", "E"]] = ["Cross Connect", "5.000", "4.800", "4.600"]
df.loc[4, ["B", "C", "D", "E"]] = ["Dismantle RAU", "3.500", "3.650", "3.450"]

st.dataframe(df, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 25px; margin-top:-10px;">
            To provide additional explanations or notes on the sheet, you can include them using an image or a text box.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:violet-badge[5. TOTAL ROW]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            You are not allowed to add a 
            <span style="font-weight: 700;">TOTAL</span> row at the bottom of the table! 
            Please refer to the example table below:
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["Desc", "Vendor A", "Vendor B", "Vendor C"]
data = [
    ["Optical Cable", "1.000", "2.000", "3.000"],
    ["Cross Connect", "5.000", "4.800", "4.600"],
    ["TOTAL", "6.000", "6.800", "7.600"],
]
df = pd.DataFrame(data, columns=columns)

def red_highlight(row):
    if row["Desc"] == "TOTAL":
        return ["color: #FF4D4D;" for _ in row]
    return [""] * len(row)

df_styled = df.style.apply(red_highlight, axis=1)

st.dataframe(df_styled, hide_index=True)

st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 20px; margin-top: -5px;">
            The table above is an 
            <span style="color: #FF69B4; font-weight: 700;">incorrect example</span> and is 
            <span style="color: #FF69B4; font-weight: 700;">not permitted</span>! 
            The total row is generated automatically during
            <span style="font-weight: 700;">MERGE DATA</span> — 
            do not add one manually, or it will be treated as part of the description and included in calculations.
        </div>
    """,
    unsafe_allow_html=True
)

st.divider()

st.markdown("#### What is Displayed?")

# Path file Excel yang sudah ada
file_paths = ["Round 1.xlsx", "Round 2.xlsx", "Round 3.xlsx", "Round 4.xlsx"]

# Buat ZIP di memory
zip_buffer = io.BytesIO()
with zipfile.ZipFile(zip_buffer, "w") as zf:
    for file_path in file_paths:
        zf.write(file_path, arcname=file_path.split("/")[-1])  # arcname = nama file di ZIP
zip_buffer.seek(0)

# Markdown teks
st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 5px; margin-top: -10px">
        You can try this menu by downloading the dummy dataset using the button below: 
    </div>
    """,
    unsafe_allow_html=True
)

@st.fragment
def release_the_balloons():
    st.balloons()

# Download button untuk file Excel
st.download_button(
    label="Dummy Dataset",
    data=zip_buffer,
    file_name="Dummy Dataset - TCO Comparison Round by Round.zip",
    mime="application/zip",
    on_click=release_the_balloons,
    type="primary",
    use_container_width=True,
)

st.markdown(
    """
    <div style="text-align: justify; font-size: 15px; margin-bottom: 20px">
        Based on this dummy dataset, the menu will produce the following results.
    </div>
    """,
    unsafe_allow_html=True
)

st.markdown("**:red-badge[1. MERGE DATA]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will merge the tables from each sheet into a single table and add
            a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">TOTAL ROW</span> for each vendor, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["ROUND", "TCO Component", "Vendor A", "Vendor B", "Vendor C"]
data = [
    ["Round 1", "Software", 12000, 13500, 11000],
    ["Round 1", "Hardware", 25000, 24000, 23000],
    ["Round 1", "TOTAL", 37000, 37500, 34000],

    ["Round 2", "Software", 9850, 10230, 9570],
    ["Round 2", "Hardware", 18020, 17590, 16980],
    ["Round 2", "TOTAL", 27870, 27820, 26550],

    ["Round 3", "Software", 14530, 15210, 13960],
    ["Round 3", "Hardware", 31080, 29840, 28590],
    ["Round 3", "TOTAL", 45610, 45050, 42550],

    ["Round 4", "Software", 13420, 12090, 11560],
    ["Round 4", "Hardware", 19570, 27840, 26510],
    ["Round 4", "TOTAL", 32990, 39930, 38070],
]
df_merge = pd.DataFrame(data, columns=columns)

num_cols = ["Vendor A", "Vendor B", "Vendor C"]
df_merge_styled = (
    df_merge.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_merge_styled, hide_index=True)

st.write("")
st.markdown("**:orange-badge[2. COST SUMMARY]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            After merging the data, the system will transpose the vendor columns into a single column and add a 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">PRICE</span>
            column as the final column in the table.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["ROUND", "VENDOR", "TCO Component", "PRICE"]
data = [
    ["Round 1", "Vendor A", "Hardware", 25000],
    ["Round 1", "Vendor A", "Software", 12000],
    ["Round 1", "Vendor A", "TOTAL", 37000],

    ["Round 1", "Vendor B", "Hardware", 24000],
    ["Round 1", "Vendor B", "Software", 13500],
    ["Round 1", "Vendor B", "TOTAL", 37500],

    ["Round 1", "Vendor C", "Hardware", 23000],
    ["Round 1", "Vendor C", "Software", 11000],
    ["Round 1", "Vendor C", "TOTAL", 34000],

    ["Round 2", "Vendor A", "Hardware", 18020],
    ["Round 2", "Vendor A", "Software", 9850],
    ["Round 2", "Vendor A", "TOTAL", 27870],

    ["Round 2", "Vendor B", "Hardware", 17590],
    ["Round 2", "Vendor B", "Software", 10230],
    ["Round 2", "Vendor B", "TOTAL", 27820],

    ["Round 2", "Vendor C", "Hardware", 16980],
    ["Round 2", "Vendor C", "Software", 9570],
    ["Round 2", "Vendor C", "TOTAL", 26550],

    ["Round 3", "Vendor A", "Hardware", 31080],
    ["Round 3", "Vendor A", "Software", 14530],
    ["Round 3", "Vendor A", "TOTAL", 45610],

    ["Round 3", "Vendor B", "Hardware", 29840],
    ["Round 3", "Vendor B", "Software", 15210],
    ["Round 3", "Vendor B", "TOTAL", 45050],

    ["Round 3", "Vendor C", "Hardware", 28590],
    ["Round 3", "Vendor C", "Software", 13960],
    ["Round 3", "Vendor C", "TOTAL", 42550],

    ["Round 4", "Vendor A", "Hardware", 19570],
    ["Round 4", "Vendor A", "Software", 13420],
    ["Round 4", "Vendor A", "TOTAL", 32990],

    ["Round 4", "Vendor B", "Hardware", 27840],
    ["Round 4", "Vendor B", "Software", 12090],
    ["Round 4", "Vendor B", "TOTAL", 39930],

    ["Round 4", "Vendor C", "Hardware", 26510],
    ["Round 4", "Vendor C", "Software", 11560],
    ["Round 4", "Vendor C", "TOTAL", 38070],
]
df_summary = pd.DataFrame(data, columns=columns)

num_cols = ["PRICE"]
df_summary_styled = (
    df_summary.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)

st.dataframe(df_summary_styled, hide_index=True)

st.write("")
st.markdown("**:yellow-badge[3. PIVOT TABLE]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system will generate a pivot table by rearranging the table structure horizontally
            based on each 
            <span style="color: #FF69B4; font-weight: 700;">'TCO Component'</span>, as shown below.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["TCO Component", "Vendor A Round 1", "Vendor A Round 2", "Vendor A Round 3", "Vendor A Round 4", "Vendor B Round 1", "Vendor B Round 2", "Vendor B Round 3", "Vendor B Round 4", "Vendor C Round 1", "Vendor C Round 2", "Vendor C Round 3", "Vendor C Round 4"]
data = [
    ["Software", 25000,18020,31080,19570,24000,17590,29840,27840,23000,16980,28590,26510],
    ["Hardware", 12000,9850,14530,13420,13500,10230,15210,12090,11000,9570,13960,11560],
    ["TOTAL", 37000,27870,45610,32990,37500,27820,45050,39930,34000,26550,42550,38070]
]
df_pivot = pd.DataFrame(data, columns=columns)

num_cols = ["Vendor A Round 1", "Vendor A Round 2", "Vendor A Round 3", "Vendor A Round 4", "Vendor B Round 1", "Vendor B Round 2", "Vendor B Round 3", "Vendor B Round 4", "Vendor C Round 1", "Vendor C Round 2", "Vendor C Round 3", "Vendor C Round 4"]
df_pivot_styled = (
    df_pivot.style
    .format({col: format_rupiah for col in num_cols})
    .apply(highlight_total, axis=1)
)
st.dataframe(df_pivot_styled, hide_index=True)

st.write("")
st.markdown("**:green-badge[4. BID & PRICE ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu also displays an analysis table that provides a comprehensive overview of the pricing structure 
            submitted by each vendor, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

st.markdown(
    """
    <div style="text-align:left; margin-bottom: 8px">
        <span style="background:#C6EFCE; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">1st Lowest</span>
        &nbsp;
        <span style="background:#FFEB9C; padding:2px 8px; border-radius:6px; font-weight:600; font-size: 0.75rem; color: black">2nd Lowest</span>
    </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["ROUND", "TCO Component", "Vendor A", "Vendor B", "Vendor C", "1st Lowest", "1st Vendor", "2nd Lowest", "2nd Vendor", "Gap 1 to 2 (%)", "Median Price", "Vendor A to Median (%)", "Vendor B to Median (%)", "Vendor C to Median (%)"]
data = [
    ["Round 1", "Software", 12000,13500,11000,11000,"Vendor C",12000,"Vendor A", 9.1, 12000, 0, 12.5, -8.3],
    ["Round 1", "Hardware", 25000,24000,23000,23000,"Vendor C",24000,"Vendor B", 4.3, 24000, 4.2, 0, -4.2],

    ["Round 2", "Software", 9850,10230,9570,9570,"Vendor C",9850,"Vendor A", 2.9, 9850, 0, 3.9, -2.8],
    ["Round 2", "Hardware", 18020,17590,16980,16980,"Vendor C",17590,"Vendor B", 3.6, 17590, 2.4, 0, -3.5],

    ["Round 3", "Software", 14530,15210,13960,13960,"Vendor C",14530,"Vendor A", 4.1, 14530, 0, 4.7, -3.9],
    ["Round 3", "Hardware", 31080,29840,28590,28590,"Vendor C",29840,"Vendor B", 4.4, 29840, 4.2, 0, -4.2],

    ["Round 4", "Software", 13420,12090,11560,11560,"Vendor C",12090,"Vendor B", 4.6, 12090, 11, 0, -4.4],
    ["Round 4", "Hardware", 19570,27840,26510,19570,"Vendor A",26510,"Vendor C", 35.5, 26510, -26.2, 5, 0],
]
df_analysis = pd.DataFrame(data, columns=columns)

num_cols = ["Vendor A", "Vendor B", "Vendor C", "1st Lowest", "2nd Lowest", "Median Price"]
format_dic = {col: format_rupiah for col in num_cols}
format_dic.update({"Gap 1 to 2 (%)": "{:.1f}%"})

vendor_cols = ["Vendor A", "Vendor B", "Vendor C"]
for v in vendor_cols:
    format_dic[f"{v} to Median (%)"] = "{:+.1f}%"

df_analysis_styled = (
    df_analysis.style
    .format(format_dic)
    .apply(lambda row: highlight_1st_2nd(row, df_analysis.columns), axis=1)
)

st.dataframe(df_analysis_styled, hide_index=True)

st.write("")
st.markdown("**:blue-badge[5. PRICE MOVEMENT ANALYSIS]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            The system also generates a price analysis table to compare price 
            <span style="color: #FF69B4">decreases</span> or 
            <span style="color: #FF69B4">increases</span> 
            across each round, as follows.
        </div>
    """,
    unsafe_allow_html=True
)

# DataFrame
columns = ["VENDOR", "TCO Component", "Round 1", "Round 2", "Round 3", "Round 4", "PRICE REDUCTION (VALUE)", "PRICE REDUCTION (%)", "PRICE TREND", "STANDARD DEVIATION", "PRICE STABILITY INDEX (%)"]
data = [
    ["Vendor A", "Hardware", 25000,18020,31080,19570,5430,21.7,"Fluctuating",5127.2428,55.8],
    ["Vendor A", "Software", 12000,9850,14530,13420,-1420,-11.8,"Fluctuating",1748.5565,37.6],
    ["Vendor A", "TOTAL", 37000,27870,45610,32990,"","","","",""],

    ["Vendor B", "Hardware", 24000,17590,29840,27840,-3840,-16,"Fluctuating",4670.8156,49.4],
    ["Vendor B", "Software", 13500,10230,15210,12090,1410,10.4,"Fluctuating",1830.292,39],
    ["Vendor B", "TOTAL", 37500,27820,45050,39930,"","","","",""],

    ["Vendor C", "Hardware", 23000,16980,28590,26510,-3510,-15.3,"Fluctuating",4399.9148,48.8],
    ["Vendor C", "Software", 11000,9570,13960,11560,-560,-5.1,"Fluctuating",1583.3568,38.1],
    ["Vendor C", "TOTAL", 34000,26550,42550,38070,"","","","",""],
]

df_pmove = pd.DataFrame(data, columns=columns)
df_pmove = df_pmove.map(lambda x: None if x == "" else x)

num_cols = ["Round 1", "Round 2", "Round 3", "Round 4", "PRICE REDUCTION (VALUE)", "STANDARD DEVIATION"]
format_dict = {col: format_rupiah for col in num_cols}
format_dict.update({
    "PRICE REDUCTION (%)": "{:+.1f}%",
    "PRICE STABILITY INDEX (%)": "{:.1f}%"
})

df_pmove_styled = (
    df_pmove.style
    .format(format_dict)
    .apply(highlight_total, axis=1)
)

st.dataframe(df_pmove_styled, hide_index=True)

st.write("")
st.markdown("**:violet-badge[6. VISUALIZATION]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            This menu displays visualizations focusing on two key aspects: 
            <span style="background: #FF5E5E; padding:1px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Winning Performance</span> and 
            <span style="background: #FF00AA; padding:2px 4px; border-radius:6px; font-weight:600; font-size: 13px; color: black">Price Trend</span>, 
            each presented in its own tab.
        </div>
    """,
    unsafe_allow_html=True
)

tab1, tab2 = st.tabs(["Winning Performance", "Price Trend"])

with tab1:
    st.image("assets/1.png")
    with st.expander("See explanation"):
            st.caption('''
                The visualization above shows the number of wins each vendor
                achieves in every tender round. A win is counted based on which
                vendor becomes the best bidder **(1st Vendor)** for each scope.
                     
                **💡 How to interpret the chart**
                     
                - High Wins Value  
                     Vendor is highly competitive in that round and wins more scopes
                     than others.  
                - Increasing Wins Across Rounds  
                     Indicates improving perfomance or more competitive pricing in later 
                     rounds.  
                - Decreasing Wins Across Rounds  
                     Shows declining competitiveness, with the vendor losing more scopes
                     compared the previous rounds.  
                - Zero Wins in a Round  
                     Vendor did not win any scope in that round, indicating weak competitiveness
                     for that stage.
            ''')

with tab2:
    st.image("assets/2.png")
    with st.expander("See explanation"):
            st.caption('''
                The chart above shows the number of occurrences of each **Price 
                Trend** for every vendor based on the pivoted tender data.
                     
                **💡 How to interpret the chart**
                     
                - No Change  
                     The vendor's price remains stable across all rounds or periods.
                - Consistently Down  
                     The vendor's price decreases continuously from one round to the next.
                - Consistently Up  
                     The vendor's price increases in every subsequent round.
                - Fluctuating  
                     The vendor's price moves up and down across the rounds.
            ''')
    
st.write("")
st.markdown("**:gray-badge[7. SUPER BUTTON]**")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            Lastly, there is a <span style="background:#FFCB09; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">Super Button</span> feature where all dataframes generated by the system 
            can be downloaded as a single file with multiple sheets. You can also customize the order of the sheets.
            The interface looks more or less like this.
        </div>
    """,
    unsafe_allow_html=True
)

dataframes = {
    "Merge Data": df_merge,
    "Cost Summary": df_summary,
    "Pivot Table": df_pivot,
    "Bid & Price Analysis": df_analysis,
    "Price Movement Analysis": df_pmove,
}

# Tampilkan multiselect
selected_sheets = st.multiselect(
    "Select sheets to download in a single Excel file:",
    options=list(dataframes.keys()),
    default=list(dataframes.keys())  # default semua dipilih
)

# Fungsi "Super Button" & Formatting
def generate_multi_sheet_excel(selected_sheets, df_dict):

    output = BytesIO()

    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        for sheet in selected_sheets:
            df_raw = df_dict[sheet].copy()

            # ===== COERCE NUMERIC SAFELY =====
            df = df_raw.copy()
            numeric_cols = []

            for col in df.columns:
                coerced = pd.to_numeric(df[col], errors="coerce")
                if coerced.notna().any():
                    df[col] = coerced
                    numeric_cols.append(col)

            pct_cols = [c for c in df.columns if "%" in c]

            df.to_excel(writer, index=False, sheet_name=sheet)
            workbook  = writer.book
            worksheet = writer.sheets[sheet]

            # ===== FORMAT =====
            fmt_rupiah = workbook.add_format({"num_format": "#,##0"})
            fmt_pct    = workbook.add_format({'num_format': '#,##0.0"%"'})

            fmt_total = workbook.add_format({
                "bold": True,
                "bg_color": "#D9EAD3",
                "font_color": "#1A5E20",
                "num_format": "#,##0"
            })

            fmt_first = workbook.add_format({
                "bg_color": "#C6EFCE",
                "num_format": "#,##0"
            })

            fmt_second = workbook.add_format({
                "bg_color": "#FFEB9C",
                "num_format": "#,##0"
            })

            # ===== COLUMN FORMAT =====
            for col_idx, col_name in enumerate(df.columns):
                if col_name in numeric_cols:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_rupiah)
                if col_name in pct_cols:
                    worksheet.set_column(col_idx, col_idx, 15, fmt_pct)

            # ===== LOOP DATA =====
            for row_idx, row in enumerate(df.itertuples(index=False), start=1):

                is_total_row = any(
                    isinstance(x, str) and x.strip().upper() == "TOTAL"
                    for x in row
                    if pd.notna(x)
                )

                # Bid & Price vendor index
                first_idx = second_idx = None
                if sheet == "Bid & Price Analysis":
                    first_vendor  = row[df.columns.get_loc("1st Vendor")]
                    second_vendor = row[df.columns.get_loc("2nd Vendor")]

                    if first_vendor in numeric_cols:
                        first_idx = df.columns.get_loc(first_vendor)
                    if second_vendor in numeric_cols:
                        second_idx = df.columns.get_loc(second_vendor)

                for col_idx, col_name in enumerate(df.columns):
                    value = row[col_idx]
                    fmt = None

                    # ===== PICK FORMAT =====
                    if sheet == "Bid & Price Analysis":
                        if col_idx == first_idx:
                            fmt = fmt_first
                        elif col_idx == second_idx:
                            fmt = fmt_second
                    elif is_total_row:
                        fmt = fmt_total

                    # ===== WRITE CELL (TYPE SAFE) =====
                    if pd.isna(value) or (isinstance(value, float) and np.isinf(value)):
                        worksheet.write_blank(row_idx, col_idx, None, fmt)

                    elif col_name in pct_cols:
                        worksheet.write_number(
                            row_idx, col_idx, value, fmt or fmt_pct
                        )

                    elif col_name in numeric_cols:
                        worksheet.write_number(
                            row_idx, col_idx, value, fmt or fmt_rupiah
                        )

                    else:
                        worksheet.write(row_idx, col_idx, value, fmt)

            # ===== AUTOFIT =====
            for i, col in enumerate(df.columns):
                worksheet.set_column(
                    i, i,
                    max(len(str(col)), df[col].astype(str).map(len).max()) + 2
                )

    output.seek(0)
    return output.getvalue()

# ---- DOWNLOAD BUTTON ----
if selected_sheets:
    excel_bytes = generate_multi_sheet_excel(selected_sheets, dataframes)

    st.download_button(
        label="Download",
        data=excel_bytes,
        file_name="Super Botton - TCO Comparison Round by Round.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary",
        use_container_width=True,
    )

st.write("")
st.divider()

st.markdown("#### Video Tutorial")
st.markdown(
    """
        <div style="text-align: justify; font-size: 15px; margin-bottom: 10px; margin-top:-10px;">
            I have also included a video tutorial, which you can access through the 
            <span style="background:#FF0000; padding:2px 4px; border-radius:6px; font-weight:600; 
            font-size: 0.75rem; color: black">YouTube</span> link below.
        </div>
    """,
    unsafe_allow_html=True
)

st.video("https://youtu.be/QcJe9ZrD-Bo?si=Ob07DiDb3_95XKiB")