import os
import tempfile

import streamlit as st

from fill_deck import generate_ppt

st.set_page_config(page_title="Fill Deck Generator", layout="wide")

st.title("📊 Fill Deck Generator")
st.caption(
    "Upload a RefractOne-style Excel workbook and a PPTX template, "
    "and get back a filled-in presentation."
)

excel = st.file_uploader("Upload Excel Workbook", type=["xlsm", "xlsx"])
ppt = st.file_uploader("Upload PPT Template", type=["pptx"])
company = st.text_input("Company Name")

if st.button("Generate Presentation", type="primary"):

    if not (excel and ppt and company):
        st.warning("Please upload both files and enter a company name before generating.")
    else:
        with st.spinner("Filling deck..."):
            try:
                with tempfile.TemporaryDirectory() as temp:
                    excel_path = os.path.join(temp, excel.name)
                    ppt_path = os.path.join(temp, ppt.name)
                    output_path = os.path.join(temp, "Filled_Deck.pptx")

                    with open(excel_path, "wb") as f:
                        f.write(excel.getbuffer())
                    with open(ppt_path, "wb") as f:
                        f.write(ppt.getbuffer())

                    stats = generate_ppt(excel_path, ppt_path, company, output_path)

                    # Read the finished deck into memory before the temp dir
                    # is cleaned up, so the download button works afterwards.
                    with open(output_path, "rb") as f:
                        deck_bytes = f.read()

            except KeyError as e:
                st.error(
                    f"The Excel workbook is missing an expected sheet or column: {e}. "
                    "Please check it follows the RefractOne schema: 'Overview', 'IT Vendor' "
                    "and 'Business Themes' are required; 'Tech Themes', 'Key Triggers', "
                    "'Sustainability & ESG Themes' and 'Challenges & Growth Opportunities' "
                    "are optional."
                )
            except Exception as e:
                st.error(f"Something went wrong while generating the deck: {e}")
            else:
                st.success("Presentation generated!")
                st.download_button(
                    "⬇ Download PPT",
                    deck_bytes,
                    file_name=f"{company.replace(' ', '_')}_Filled_Deck.pptx",
                    mime="application/vnd.openxmlformats-officedocument.presentationml.presentation",
                )

                if stats.get("unplaced_vendors"):
                    st.warning(
                        "These vendors have a Level 1 Category that is neither "
                        "'Software' nor 'Services', so they were left off the vendor "
                        "slides: " + ", ".join(stats["unplaced_vendors"])
                    )

                with st.expander("Details"):
                    st.write(
                        f"- IT Vendors — Software: {stats['n_software']} slide(s)\n"
                        f"- IT Vendors — ICT Services: {stats['n_ict']} slide(s)\n"
                        f"- Business Themes: {stats['n_business']} slide(s) "
                        f"for {stats['n_business_themes']} theme(s)\n"
                        f"- Challenges & Growth Prospects: {stats['n_challenges']} slide(s)\n"
                        f"- Sustainability Themes: {stats['n_esg']} slide(s) "
                        f"for {stats['n_esg_themes']} theme(s)\n"
                        f"- Technology Themes: {stats['n_tech']} slide(s) "
                        f"for {stats['n_tech_themes']} theme(s)\n"
                        f"- Key Developments: {stats['n_keydev']} slide(s) "
                        f"for {stats['n_events']} event(s)"
                    )
