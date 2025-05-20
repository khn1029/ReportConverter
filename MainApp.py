import streamlit as st
import io
import importlib
import basic_functions
from basic_functions import *
from docx import Document
from PIL import Image
import pandas as pd

importlib.reload(basic_functions)

st.set_page_config(layout="wide")
st.title("Report Converter")

uploaded_file = st.file_uploader("Upload a GWP or Appendix .docx file", type="docx")

doc_type = ['Geological Well Proposal (GWP) Main File','GWP Appendix']
selected_doc_type = st.selectbox("Document is", doc_type)
status_placeholder = st.empty()

@st.cache_resource
def load_docx(file):
    return Document(file)

@st.cache_data
def process_docx(_doc,phrase,):
    object_list = load_object_list(_doc,phrase)
    return object_list

def create_excel_report(object_list,progress_callback=None):
    output = io.BytesIO()
    export_report_to_excel(object_list, output,progress_callback)
    output.seek(0)
    return output

if (st.button("Run")) and (uploaded_file is not None):

    st.success(f"{uploaded_file.name} DOCX file uploaded successfully.")
    doc = load_docx(uploaded_file)
    st.session_state["doc_name"] = uploaded_file.name
    ### select phrase
    if selected_doc_type == doc_type[0]:  # GWP
        phrase = 'executive summary'
    else:  # Appendix
        phrase = 'appendix'

    st.info("Extracting contents...")
    object_list = process_docx(doc,phrase)
    st.session_state["object_list"] = object_list  # store to streamlit

    st.success("Report generated successfully!")
    st.session_state["run_complete"] = True  # ✅ mark as run
# Dropdown
if st.session_state.get("run_complete"):
    object_list=st.session_state["object_list"]
    labels = [f"{i}: {obj.get('label', 'No Label')} - {obj.get('title', 'No Title')}" for i, obj in enumerate(object_list)]
    selected_label = st.selectbox("Select an object to view", labels)
    selected_idx = int(selected_label.split(":")[0])
    selected_obj = object_list[selected_idx]

    st.subheader(selected_obj.get("title", "Untitled"))
    st.caption(selected_obj.get("label", ""))

    data = selected_obj.get("data")

    if isinstance(data, pd.DataFrame):
        show_dataframe_as_table(data)
    else:
        try:
            image = Image.open(io.BytesIO(data)).copy()
            st.image(image, caption=selected_obj.get("title", "Image"), use_column_width=True)
        except:
            st.warning("No data to display or unsupported format.")

# Initialize the state
if "create_excel_clicked" not in st.session_state:
    st.session_state["create_excel_clicked"] = False

# CREATE AND DOWNLOAD EXCEL
if st.button("Create Excel report (each figure/table takes a sheet)"):
    if "object_list" in st.session_state:
        st.session_state["create_excel_clicked"] = True  # ✅ Set flag
        progress_text = st.empty()
        progress_bar = st.progress(0)

        def progress_callback(done, total):
            percent = done / total
            progress_bar.progress(percent)
            progress_text.write(f"✅ Completed {done} of {total} sheets")

        with st.spinner("Creating Excel file..."):
            output = create_excel_report(
                st.session_state["object_list"],
                progress_callback # call a function
            )
        st.session_state["excel_output"] = output
        progress_bar.empty()
        progress_text.empty()

        # st.download_button(
        #     label="📥 Download Excel report file",
        #     data=output,
        #     file_name=f"{st.session_state['doc_name']}_excel_report.xlsx",
        #     mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        # )
    else:
        st.warning("You must run the document processing step first.")

# Show download button only after the first button is clicked and output is ready
if st.session_state.get("create_excel_clicked") and ("excel_output" in st.session_state):
    st.download_button(
        label="📥 Download Excel report file",
        data=st.session_state["excel_output"],
        file_name=f"{st.session_state['doc_name']}_excel_report.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        key="excel_download_button"
    )