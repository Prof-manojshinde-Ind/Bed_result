import streamlit as st
import pdfplumber
import re
import pandas as pd
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Alignment
import time
import matplotlib.pyplot as plt

# -----------------------------
# Classification Function
# -----------------------------
def classify_student(records, percentage):
    subs_101_107 = [m for s, m in records if s.startswith("BED10") and int(s[-3:]) <= 107]
    subs_108_112 = [m for s, m in records if s.startswith("BED10") and 108 <= int(s[-3:]) <= 112]
    subs_201_205 = [m for s, m in records if s.startswith("BED20") and int(s[-3:]) <= 205]
    subs_206_212 = [m for s, m in records if s.startswith("BED20") and 206 <= int(s[-3:]) <= 212]

    # ---- First Year ----
    if subs_101_107:
        if any(m == 0 for m in subs_101_107): return "Fail"
        if sum(m < 50 for m in subs_101_107) > 3: return "Fail"
        if 1 <= sum(m < 50 for m in subs_101_107) <= 3: return "ATKT"
        if any(m == 0 for m in subs_108_112): return "Fail in Internal"

    # ---- Second Year ----
    if subs_201_205 and any(m < 50 for m in subs_201_205): return "Fail"
    if subs_206_212 and any(m == 0 for m in subs_206_212): return "Fail in Internal"

    # ---- Percentage based ----
    if percentage >= 80: return "Distinction"
    elif percentage >= 65: return "First Class"
    elif percentage >= 55: return "Second Class"
    elif percentage >= 50: return "Pass"
    return "Fail"

# -----------------------------
# PDF Parser
# -----------------------------
def parse_pdf(file):
    students_data, current_student = {}, None
    with pdfplumber.open(file) as pdf:
        for page in pdf.pages:
            lines = (page.extract_text() or "").split("\n")
            for line in lines:
                if line.startswith("PRN:"):
                    match = re.search(r"SEAT NO\.\:\s*(\d+)\s+NAME:\s*(.*?)\s+Mother", line)
                    if match:
                        seat, name = match.group(1).strip(), match.group(2).strip()
                        current_student = (seat, name)
                        students_data[current_student] = []
                match = re.match(r"(BED\s+\d{3}(?:-\d{2})?)\s+.*\s(\d{2,3})\s+\d+\s+\d+\s+\w+", line)
                if match and current_student:
                    subj, marks = match.group(1).replace(" ", ""), int(match.group(2))
                    students_data[current_student].append([subj, marks])
    return students_data

# -----------------------------
# Excel Builder
# -----------------------------
def build_excel(students_data):
    wb = Workbook()
    ws_all = wb.active
    ws_all.title = "All_Students"
    ws_all.append(["SEAT NO", "NAME", "Subject", "Marks", "Total", "Percentage", "Class"])

    row, summary_fy, summary_sy, all_cls, fy_cls, sy_cls = 2, [], [], [], [], []
    for (seat, name), records in students_data.items():
        start, total, last_sub = row, 0, None
        for subj, marks in records:
            total += marks
            last_sub = subj
            ws_all.append([seat if row == start else None,
                           name if row == start else None,
                           subj, marks])
            row += 1
        end = row - 1
        perc, classification = None, None
        if last_sub.endswith("112"):
            perc, classification = (total / 1000) * 100, classify_student(records, (total / 1000) * 100)
            summary_fy.append([seat, name, total, round(perc, 2), classification]); fy_cls.append(classification)
        elif last_sub.endswith("212"):
            perc, classification = (total / 2000) * 100, classify_student(records, (total / 2000) * 100)
            summary_sy.append([seat, name, total, round(perc, 2), classification]); sy_cls.append(classification)
        all_cls.append(classification)
        ws_all.cell(end, 5, total)
        if perc: ws_all.cell(end, 6, round(perc, 2)).number_format = "0.00"
        ws_all.cell(end, 7, classification)
        ws_all.merge_cells(start_row=start, start_column=1, end_row=end, end_column=1)
        ws_all.merge_cells(start_row=start, start_column=2, end_row=end, end_column=2)
        for col in (1, 2): ws_all.cell(start, col).alignment = Alignment(horizontal="center", vertical="center")

    # First & Second Year sheets
    def make_sheet(title, data):
        ws = wb.create_sheet(title)
        ws.append(["SEAT NO", "NAME", "Total", "Percentage", "Class"])
        for r in sorted(data, key=lambda x: x[3], reverse=True)[:5]:
            ws.append(r); ws.cell(ws.max_row, 4).number_format = "0.00"
    make_sheet("First Year", summary_fy)
    make_sheet("Second Year", summary_sy)

    out = BytesIO(); wb.save(out); out.seek(0)
    return out, summary_fy, summary_sy, all_cls, fy_cls, sy_cls

# -----------------------------
# Matplotlib Plotting
# -----------------------------
def plot_class_distribution(counts, title):
    df = counts.reset_index(); df.columns = ["Class", "Count"]
    fig, ax = plt.subplots(figsize=(6,4))  # ✅ medium size
    bars = ax.bar(df["Class"], df["Count"])
    ax.set_title(title)
    ax.set_ylabel("Count")
    for bar in bars:
        yval = bar.get_height()
        ax.text(bar.get_x()+bar.get_width()/2, yval+0.2, int(yval),
                ha="center", va="bottom")
    st.pyplot(fig)

    # Table with serial number starting from 1
    df.index = df.index + 1
    df.index.name = "S.No"
    st.table(df)

def plot_side_by_side(fy_counts, sy_counts):
    df = pd.DataFrame({"First Year": fy_counts, "Second Year": sy_counts}).fillna(0).reset_index().rename(columns={"index":"Class"})
    fig, ax = plt.subplots(figsize=(7,4))  # ✅ medium size
    width = 0.35
    x = range(len(df))
    bars1 = ax.bar([i-width/2 for i in x], df["First Year"], width, label="First Year")
    bars2 = ax.bar([i+width/2 for i in x], df["Second Year"], width, label="Second Year")
    ax.set_xticks(x); ax.set_xticklabels(df["Class"])
    ax.set_title("FY vs SY — Class Distribution")
    ax.legend()
    for bars in [bars1, bars2]:
        for bar in bars:
            yval = bar.get_height()
            ax.text(bar.get_x()+bar.get_width()/2, yval+0.2, int(yval),
                    ha="center", va="bottom")
    st.pyplot(fig)

    df.index = df.index + 1
    df.index.name = "S.No"
    st.table(df)

# -----------------------------
# Streamlit UI
# -----------------------------
st.title("📊 Student Result Processor")
st.write("Upload Result Summary PDF to extract and analyze results.")
st.write("Code by Prof.Manoj Shinde (prof.manojshinde@gmail.com,+91 7020799486)")

uploaded = st.file_uploader("Upload PDF File", type=["pdf"])
if uploaded:
    start = time.time()
    with st.spinner("⏳ Processing..."):
        data = parse_pdf(uploaded)
    if not data:
        st.error("❌ No student data found in this PDF.")
    else:
        excel, sum_fy, sum_sy, all_cls, fy_cls, sy_cls = build_excel(data)
        st.info(f"⏱ Completed in {time.time()-start:.2f} seconds.")

        # ✅ DataFrames with Top 5 by Percentage
        df_fy = pd.DataFrame(sum_fy, columns=["SEAT NO","NAME","Total","Percentage","Class"])
        df_sy = pd.DataFrame(sum_sy, columns=["SEAT NO","NAME","Total","Percentage","Class"])
        df_fy = df_fy.sort_values(by="Percentage", ascending=False).head(5); df_fy.index = range(1, len(df_fy)+1)
        df_sy = df_sy.sort_values(by="Percentage", ascending=False).head(5); df_sy.index = range(1, len(df_sy)+1)

        st.subheader("🏆 First Year - Top 5 Students")
        st.table(df_fy)

        st.subheader("🏆 Second Year - Top 5 Students")
        st.table(df_sy)

        st.subheader("📊 Class Distribution (All Students)")
        plot_class_distribution(pd.Series(all_cls).value_counts().sort_index(), "All Students")

        st.subheader("📊 First Year Distribution")
        plot_class_distribution(pd.Series(fy_cls).value_counts().sort_index(), "First Year")

        st.subheader("📊 Second Year Distribution")
        plot_class_distribution(pd.Series(sy_cls).value_counts().sort_index(), "Second Year")

        st.subheader("📊 Side-by-Side (FY vs SY)")
        plot_side_by_side(pd.Series(fy_cls).value_counts().sort_index(), pd.Series(sy_cls).value_counts().sort_index())

        st.download_button("📥 Download Excel", data=excel,
            file_name="all_final_result.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
