import numpy as np
import pandas as pd
import pdfkit
from fitz import fitz, Rect
import re

tutor_initials_template = re.compile("\\b[a-zA-Z]")
filename_template = re.compile(r'([A-Z]|[a-z]).*[.]\d{4}[.]\d{2}[.]\d{2} [A-Z](.*)[A-Z][_]\d')

def filename_generator(df, index):
    tutor_initials = ""
    tutor_name = df.iloc[index]["Tutor Name:"]
    teacher = df.iloc[index]["Teacher Name:"]
    date = str(df.iloc[index]["The date when the session took place"])
    id = str(df.iloc[index]["ID"])

    #teacher last name generator
    for word in re.finditer(tutor_initials_template, tutor_name):
        initial = word.group(0)
        tutor_initials = tutor_initials + initial

    teacher_name_lst = list(teacher.split(" "))
    teacher_lastname = teacher_name_lst[-1]

    # date generator
    date_lst1 = list(date.split(" "))
    date_only = date_lst1[0]
    date_lst2 = list(date_only.split("-"))
    date = '.'.join(date_lst2)

    filename = teacher_lastname + "." + date + " " + tutor_initials + "_" + id

    # catching errors
    match = re.search(filename_template, filename)
    if not match:
        print(f"ERROR: please check report {id}. Teacher: {teacher_lastname}")

    print(filename)
    return filename


def dfs_to_pdfs(df, id):
    # rect = fitz.Rect(0, 0, 100, 100)
    for ind in range(id, len(df.index)):
        intermediate_series = df.iloc[ind, :]
        intermediate_df = intermediate_series.to_frame()
        intermediate_df = intermediate_df.fillna(" ")
        filename = filename_generator(df, ind)
        intermediate_df.to_html(f"{filename}.html")
        pdfkit.from_file(f"{filename}.html", f"{filename}.pdf")

        """doc = fitz.open(f"{filename}.pdf")
        for page in doc:
            page.insert_image(rect, filename="WritingCenterLogo.png")
        doc.saveIncr()"""


def check_missing_values(df):
    nan_cols = df.columns[df.isna().any()].tolist()
    nan_cols.remove("If the session was online, did your tutee confirm that A) they were not in a class and B) they were not asking for tutoring about a test or exam.")
    for col in nan_cols:
        rows = df.loc[df[col].isnull()]
        for index in range(0, len(rows)):
            id = str(rows.iloc[index]["ID"])
            print(f"Missing elements in report #{id}: {col}")

    print(f"\n")


# running the code
dataframe = pd.read_excel(r'/Users/sabrinadu/Documents/JAC/English/Tutoring/TutorReportForm.xlsx')
report_to_start_with = 19
check_missing_values(dataframe)
dfs_to_pdfs(dataframe, report_to_start_with)
