import pandas as pd
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
from collections import ChainMap


def streamlitx():
    df = pd.read_excel("Consolidated_file.xlsx")
    st.set_page_config(
        page_title="EJK Corpora",
        page_icon="🧊",
        layout="wide",
    )
    st.cache()

    reduce_header_height_style = """
                    <style>
                        div.block-container {padding-top:2rem;}
                        cellStyle: {textAlign: 'center'}
                    </style>
                """
    st.markdown(reduce_header_height_style, unsafe_allow_html=True)

    tab1, tab2 = st.tabs(["Data", "Summary"])
    with tab1:
        col1, col2 = st.columns([5, 2])
        with col1:
            col1.subheader("Data Summary")
            options_builder = GridOptionsBuilder.from_dataframe(df)
            options_builder.configure_column('Entry No.', displayName='', width=1, menuTabs=[], suppressMenu=True,
                                             suppressSizeToFit=True)
            options_builder.configure_column('School', width=5, animateRows='true')
            options_builder.configure_column('EssayNo.', displayName='No.', width=5, suppressMenu=True)
            options_builder.configure_column('Strand', width=5, suppressMenu=True)
            options_builder.configure_column('Section', width=3, suppressMenu=True)
            options_builder.configure_column('Listed Word Count', displayName='Count', width=4)
            options_builder.configure_column('Evaluated Word Count', hide=True)
            options_builder.configure_column('Essay', hide=True)
            options_builder.configure_selection('single')
            options_builder.configure_side_bar(filters_panel=True, columns_panel=False, defaultToolPanel="")
            grid_options = options_builder.build()

            grid_return = AgGrid(df, grid_options, theme='material', enable_enterprise_modules=True,
                                 height=450, width=700, autoSizeColumn=True, skipHeaderOnAutoSize=True,
                                 suppressSizeToFit=True)
        selected_rows = grid_return['selected_rows']

        data = dict(ChainMap(*selected_rows))
        with col2:
            col2.subheader("Click a data entry!")
            st.write("School:", data.get("School"))
            st.write("Grade:", data.get("Grade"))
            st.write("Strand and Section:", data.get("Strand"), data.get("Section"))
            st.write("Essay Number:", data.get("EssayNo."))
            st.write("Word Count:", data.get("Listed Word Count"))
            st.write(data.get("Essay"))

if __name__ == "__main__":
    streamlitx()