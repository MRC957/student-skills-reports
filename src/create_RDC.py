# -*- coding: utf-8 -*-
"""
Sébastien Van Laecke - 25/03/2024
"""

import pandas as pd
import os
import logging
import coloredlogs
from logging.handlers import RotatingFileHandler
import shutil
import win32com.client as win32
import time

from PyPDF2 import PdfMerger


if "src" in os.listdir(): src_dir = os.path.join(os.getcwd(), "src")
elif os.getcwd().endswith("src") : src_dir = os.getcwd()
else: raise Exception("'src' folder not found")

DIR_OUT = os.path.join(src_dir, "outputs")
DIR_IN = os.path.join(src_dir, "inputs")
MERGE_ONLY = False


# ------------------- create logger -------------------
logger = logging.getLogger(__name__)
logger.setLevel(logging.DEBUG)

# create console handler and set level to debug
ch = logging.StreamHandler()
ch.setLevel(logging.DEBUG)

# create file handler and set level to info
fh = RotatingFileHandler('create_RDC.log', backupCount=1, maxBytes=1e6)
fh.setLevel(logging.INFO)

# create formatter
formatter = logging.Formatter('%(asctime)s - %(name)s - %(levelname)s - %(message)s')

# add formatter to handlers
ch.setFormatter(formatter)
fh.setFormatter(formatter)

# add handlers to logger
logger.addHandler(ch)
logger.addHandler(fh)

coloredlogs.install(level='INFO', logger=logger, fmt='%(asctime)s - %(name)s - %(levelname)s - %(message)s')
# ------------------- end of create logger -------------------

class RDCProcessor():

    def __init__():
        pass

    def compute_df(self, df):

        # select the columns where the value of the row with Name == "-" is not "-"
        index_of_row_total = df.index[df["Name"] == "-"][0]
        cols_to_keep = ["Name", "First name"] + [col for col in df.columns if df.loc[index_of_row_total, col] != "-"]

        # select only the columns to keep
        df = df[cols_to_keep]

        # Select the rows containing test scores
        df_tests = df.iloc[1, 4:]
        df_tests = df_tests.astype(int).sort_values()
        self.test_new_table_start = df_tests[(df_tests // 10) != (df_tests // 10).shift(1)].index

        max_scores = pd.to_numeric(df.loc[index_of_row_total, df_tests.index])
        columns_sorted = df.columns[:3].append(df_tests.index)

        df = df[columns_sorted]

        df_tests = df.iloc[2:,3:]
        df_tests = df_tests.apply(pd.to_numeric, errors='coerce')

        df_tests[df_tests < max_scores/2] = 0
        df_tests[df_tests >= max_scores/2] = 1
        df_tests.fillna(2, inplace=True)
        df_tests = df_tests.astype(int).astype(str)

        # Merge df of names and results
        df = df.iloc[2:,:3].join(df_tests)
        return df

    def open_word(self, filename, word):

        # Open existing Word document
        doc = word.Documents.Open(filename)
        # doc_path_full = os.path.join(os.getcwd(), filename)
        # doc = word.Documents.Open(doc_path_full)
        time.sleep(0.1)

        # Close the document and quit Word
        doc.Close()
        return

    def merge_pdf(self):

        merger = PdfMerger()
        dir_pdf = os.path.join(DIR_OUT, "pdf_to_be_merged")
        if not os.path.exists(dir_pdf) :
            os.makedirs(dir_pdf)

        # List current directory to find class directories
        for classe in os.listdir(DIR_OUT):
            if os.path.isdir(os.path.join(DIR_OUT, classe)):
                
                pdf_folder_path = os.path.join(DIR_OUT, classe, "PDF")
                if os.path.isdir(pdf_folder_path):

                    # List PDF folder and merge all *.PDF in <class>.pdf
                    for file in os.listdir(pdf_folder_path):
                        if file.endswith("pdf"):
                            logger.info("Adding {} to {}.pdf ...".format(str(file), str(classe)))
                            filepath = os.path.join(pdf_folder_path, file)
                            merger.append(filepath)

                    merger.write(os.path.join(dir_pdf, "{}.pdf".format(classe)))
                    merger.close()
                    logger.info("{}.pdf created !".format(classe))
                    merger = PdfMerger()

                    # Remove the PDF folder
                    # shutil.rmtree(pdf_folder_path)

        # List current directory to find .pdf
        for file in os.listdir(dir_pdf):
            if file[-3:].lower()=="pdf":
                logger.info("Adding {} to All.pdf".format(str(file)))
                filepath = os.path.join(dir_pdf, file)
                merger.append(filepath)

        merger.write(os.path.join(DIR_OUT, "All.pdf"))
        merger.close()
        shutil.rmtree(dir_pdf)

        logger.info("FINISHED : All.pdf created !")

    def clean_folder(self):
        for item in os.listdir(DIR_OUT):
            if not item.endswith('.csv'):
                path = os.path.join(DIR_OUT, item)
                if os.path.isdir(path):
                    shutil.rmtree(path)
                else:
                    os.remove(path)
                logger.info(f"Item removed : {str(path)}")

    def execute(self):

        t0 = time.time()

        self.clean_folder()

        columns = ["Classe", "Name", "First name", "-"] + [f"test {i}" for i in range(1,11)]
        df_full = pd.read_csv(os.path.join(DIR_OUT, "Data.csv"), header=None, encoding='ANSI', names=columns, index_col=None)
        logger.debug("csv succesfully read : {}sec".format(round(time.time()-t0, 3)))

        # Open Word application
        word = win32.Dispatch('Word.Application')
        word.Visible = False

        for classe in df_full["Classe"].unique():
            
            logger.info("Processing classe {}..".format(classe))
            t1 = time.time()

            class_path = os.path.join(DIR_OUT, classe)
            if not os.path.exists(class_path) :
                os.makedirs(class_path)

            df = df_full[df_full["Classe"] == classe]
            df = self.compute_df(df)

            logger.debug("Compute df : {}sec".format(round(time.time()-t1, 3)))
            t1 = time.time()

            for file in os.listdir(DIR_IN):
                if file.endswith('.docm') and classe.lower() in file.lower():
                    template_file = file
                    break
            else:
                logger.warning(f"Word template not found for classe {classe}")
                continue


            for row in df.iterrows():
                # Retrieve row data
                row = row[1]
                if row["Name"].lower() in ["-", "nom"] : continue

                logger.info("Processing {} {}..".format(row["Name"], row["First name"]))

                # Define filename with double underscore between tables : "Name_1_Surname_1_Class1__Table1Test1_Table1Test2_...__Table2Test1_Table2Test2_...""
                filename = ""
                for k, v in row.items():
                    if k not in self.test_new_table_start:
                        filename += f"{v}_"
                    else:
                        filename += f"_{v}_"
                filename = filename[:-1]


                source_file = os.path.join(DIR_IN, template_file)
                target_file = os.path.join(DIR_OUT, row["Classe"], f"{filename}.docm")
                # source_file = os.path.join(row["Classe"], WORD_TEMPLATE_NAME)
                # target_file = os.path.join(row["Classe"], f"{filename}.docm")

                # Create Word and PDF
                shutil.copy(source_file, target_file)
                
                t2 = time.time()
                self.open_word(target_file, word)
                logger.debug("Open/close Word : {}sec".format(round(time.time()-t2, 3)))

            logger.debug("Compute of full classe : {}sec --> {}sec/student".format(round(time.time()-t1, 2), \
                round((time.time()-t1)/len(df. index), 2)))
            

        word.Quit()

        t3 = time.time()
        self.merge_pdf()
        logger.debug("Merge all PDF : {}sec".format(round(time.time()-t3, 3)))

        logger.info("end of script")


if __name__ == '__main__':
    if MERGE_ONLY:
        RDCProcessor().merge_pdf()
    else:
        RDCProcessor().execute()