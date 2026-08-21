"""
app.py
------
The web "platform". Run this with:  streamlit run app.py
It gives you a page in your browser with just one upload button. Upload
your Excel workbook (one sheet per executive - LA1, LA2, ... LAn - each
sheet holding that person's LinkedIn URL, connection count, and activity
rows), click Generate, and get back ONE PowerPoint covering every sheet.
"""

import streamlit as st
from ppt_generator import generate_pptx, read_workbook

st.set_page_config(page_title="LinkedIn Analysis PPT Generator", page_icon="📊")

st.title("📊 LinkedIn Analysis PPT Generator")
st.write(
    "Upload your Excel workbook. Every sheet that contains LinkedIn "
    "activity data (PostType, TimeAgo, Content, PostLink, Source, "
    "Headline columns, plus a LinkedIn URL and Connections value near the "
    "top) gets its own set of slides in the output - all built from the "
    "`Sample.pptx` template automatically."
)

TEMPLATE_PATH = "Sample.pptx"  # must sit next to app.py in the repo

uploaded_excel = st.file_uploader(
    "Upload your Excel workbook (.xlsx or .xlsm)", type=["xlsx", "xlsm"]
)

rows_per_slide = st.number_input(
    "Rows per slide (leave as-is unless you changed the template's table)",
    min_value=1, max_value=20, value=3, step=1,
)

if uploaded_excel is not None:
    if st.button("Generate PowerPoint"):
        with st.spinner("Building your presentation..."):
            try:
                # Quick preview of what will be included, before generating.
                uploaded_excel.seek(0)
                sheets = read_workbook(uploaded_excel)
                st.write(f"Found **{len(sheets)}** sheet(s) with data:")
                st.table([
                    {"Sheet": s["sheet_name"], "LinkedIn URL": s["linkedin_url"],
                     "Posts": len(s["rows"])}
                    for s in sheets
                ])

                uploaded_excel.seek(0)
                output = generate_pptx(
                    TEMPLATE_PATH,
                    uploaded_excel,
                    rows_per_slide=rows_per_slide,
                )
                st.success("Done! Your presentation is ready.")
                st.download_button(
                    label="⬇️ Download PPTX",
                    data=output,
                    file_name="LinkedIn_Analysis_Output.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )
            except Exception as e:
                st.error(f"Something went wrong: {e}")