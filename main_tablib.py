import tablib
import glob
import os
import re
import sys
import warnings
from datetime import datetime
import docx2txt
import pandas as pd
import pyfiglet
import streamlit as st
from st_aggrid import AgGrid, GridOptionsBuilder
from collections import ChainMap

counts = []
realcounts = []


def convert_docx():
    docxfiles = glob.glob('./input/*.docx')
    count = len(docxfiles)
    print(f"There are {count} files:")
    if not docxfiles:
        print("There are no compatible files in the input directory.")
        print("Only docx files are compatible at the moment.")
        premature()

    for docx in docxfiles:
        print("     " + docx)
        text = docx2txt.process(docx)
        textnew = re.sub(r'\n\s*\n', '\n', text)
        docxnoext = os.path.splitext(docx)[0]
        with open(docxnoext + ".txt", "w", encoding="utf-8") as text_file:
            print(textnew, file=text_file)
            print("\n" + "School:", file=text_file)
    print("Conversion complete!" + "\n\n")


def regex():
    list_of_files = glob.glob('./input/*.txt')
    if not list_of_files:
        print("There are no text files to process")
        exit()
    for file_name in list_of_files:
        fname = os.path.basename(file_name)
        print(f"\n\n\nNow working with {fname}")

        with open(file_name, encoding='utf-8') as textfile:
            filetext = textfile.read()

            dirname = os.path.dirname(__file__)
            outputpath = os.path.join(dirname)
            filenamenoext = os.path.splitext(file_name)[0]
            with open(outputpath + filenamenoext + ".json", 'w+', encoding="utf-8") as wf:
                data = tablib.Dataset()
                regtextblocs = r"""(?<=School:)([\s\S\n]+?)(?=School)"""
                blockreg = re.finditer(regtextblocs, filetext, re.MULTILINE | re.VERBOSE)
                i = 0
                for match in blockreg:
                    i += 1
                    result = match.group()

                    school = (result.split('\n', 1)[0]).replace(" ", "")
                    essayno = "not provided"
                    try:
                        essayno = int(''.join(filter(str.isdigit, result.splitlines()[1])))
                    except:
                        pass
                    grade = ("not provided")
                    try:
                        grade = int(''.join(filter(str.isdigit, result.splitlines()[5])))
                    except:
                        pass

                    strand = (result.splitlines()[3]).split(":")[-1].replace(" ", "")
                    section = (result.splitlines()[4]).split(":")[-1].replace(" ", "")
                    lwc = 0
                    try:
                        lwc = int(''.join(filter(str.isdigit, result.splitlines()[7])))
                    except:
                        print(f"        Essay match no {i} in {filenamenoext} "
                              f"wasn't provided with a word count.")
                        input("         Press enter to ignore. Please fix the problem manually."
                              "\n         Exit to terminate the program before fixing.")

                    essay = ''.join(result.splitlines(keepends=True)[8:])
                    rwc = (len(essay.split()))

                    counts.append(lwc)
                    realcounts.append(rwc)
                    data.append([school, essayno, grade, strand, section, lwc, rwc, essay])

                data.headers = ['School', 'EssayNo.', 'Grade', 'Strand', 'Section',
                                'Listed Word Count', 'Evaluated Word Count', 'Essay']
                wf.write(data.json)


def combineexcel():
    print('Consolidating data collected to a file.')
    xlsxfiles = glob.glob('./input/*.json')
    df = pd.DataFrame()
    for xlsx in xlsxfiles:
        df = df.append(pd.read_json(xlsx), ignore_index=True)
    df2 = df.reset_index(drop=True)
    df2.index = df.index.rename('Entry No.')
    df2.index += 1
    df2.to_excel("Consolidated_file.xlsx")


def sums():
    newcounts = []
    for item in counts:
        try:
            listedwc = int(item)
            newcounts.append(listedwc)
        except ValueError:
            pass

    total = sum(newcounts)
    print("The total is:")
    result = pyfiglet.figlet_format(str(total))
    print(result)
    needleft = 1000000 - total
    print(f"You need {needleft} more words.")
    realadd = [item for item in realcounts if isinstance(item, (int, float))]
    realtotal = sum(realadd)
    print(realtotal)


def premature():
    print("Process ended prematurely.")
    input("Press any key to exit")
    sys.exit()


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
            col1.subheader("A wide column with a chart")
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
            col2.subheader("A narrow column with the data")
            st.write("School:", data.get("School"))
            st.write("Grade:", data.get("Grade"))
            st.write("Strand and Section:", data.get("Strand"), data.get("Section"))
            st.write("Essay Number:", data.get("EssayNo."))
            st.write("Word Count:", data.get("Listed Word Count"))
            st.write(data.get("Essay"))



if __name__ == "__main__":
    warnings.filterwarnings(action='ignore')
    start = datetime.now()
    # convert_docx()
    regex()
    combineexcel()
    sums()
    streamlitx()

    with open("timer.txt", 'a', encoding="utf-8") as timer:
        totaltime = (datetime.now() - start)
        print(f"This process took {totaltime}")
        timer.write(str(datetime.now() - start) + "\n")

    # input("Press Enter to continue.")
