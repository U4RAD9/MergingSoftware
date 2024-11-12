import tkinter as tk
from tkinter import messagebox, filedialog
from pathlib import Path
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
import PyPDF2
import os
import re
import time
import shutil
import openpyxl
from PIL import Image
import pytesseract
import fitz
from openpyxl.styles import Font, PatternFill
from openpyxl import Workbook
import pandas as pd
from concurrent.futures import ThreadPoolExecutor
import math
# import pydicom
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from datetime import datetime

import warnings

warnings.filterwarnings("ignore", category=UserWarning, module="PyPDF2._cmap")

def merge_redcliffe_pdf_files():
    # Prompt user to select input directory
    input_directory = filedialog.askdirectory(title="Select Input Directory")

    if input_directory:
        # Prompt user to select output directory
        output_directory = filedialog.askdirectory(title="Select Output Directory")

        if output_directory:
            pdf_dir = Path(input_directory)
            pdf_output_dir = Path(output_directory)
            pdf_output_dir.mkdir(parents=True, exist_ok=True)

            # List all PDF files in the input directory
            pdf_files = list(pdf_dir.glob("*.pdf"))
            print("These are the pdf files list :", pdf_files)

            keys = set([str(file).split("\\")[-1].split("_")[0].lower() for file in pdf_files])

            for key in keys:
                xray = None
                optometry = None
                ecg_report = None
                pft = None
                audiometry = None
                vitals = None
                others = None
                blood_report = None
                smart_report = None
                # Adding this too (Himanshu/12nov24)
                vaccination_report = None
                pdf_files_3_or_less = False
                total_pdfs_less_than_3 = 0

                print("This is the key :", key)

                for file in pdf_files:
                    print("Entering the for files in pdf_files loop")
                    str_pdf_file = str(file)
                    print(str_pdf_file)
                    split_str_pdf_files = str_pdf_file.split("_")[1].lower()
                    print("This is the splitted str :", split_str_pdf_files)
                    splitted_file = str_pdf_file.rsplit("\\",1)[1].lower()
                    print("This is the splitted file :", splitted_file)
                    try:
                        print("Inside the try block.")
                        # if splitted_file.startswith(key):
                        #     print("Starts with key")
                        # if split_str_pdf_files.endswith(key):
                        if splitted_file.startswith(key):
                            print("Starts with key")
                            first_page_text = ''
                            second_page_text = ''
                            print("Before opening the pdf file page.")
                            pdf_reader = PdfReader(open(file, "rb"))
                            print("After reading the pdf.")
                            if len(pdf_reader.pages) == 1:
                                print("if the pdf reader length is equal to 1")
                                first_page = pdf_reader.pages[0]
                                first_page_text = first_page.extract_text()
                            else:
                                print("this is the else statement after calculating the pdf length.")
                                second_page = pdf_reader.pages[1]
                                second_page_text = second_page.extract_text()

                            print("This is the first page text :")
                            print(first_page_text)
                            print("End of first page text")
                            print("This is the second page text:")
                            print(second_page_text)
                            print("End of second page text")

                            if "X-RAY" in first_page_text or "X-RAY" in second_page_text:
                                xray = file
                                print("This is an xray file.")
                            elif "OPTOMETRY" in first_page_text:
                                optometry = file
                                print("This is an opto file.")
                            elif "ECG" in second_page_text or "Acquired on" in second_page_text:
                                ecg_report = file
                                print("This is an ecg file.")
                            elif "RECORDERS & MEDICARE SYSTEMS" in first_page_text:
                                pft = file
                                print("This is a pft file.")
                            elif "VITALS" in first_page_text:
                                vitals = file
                                print("This is a vitals file.")
                            elif 'left ear' in first_page_text:
                                audiometry = file
                                print("This is a audio file.")
                            elif 'RBC Count' in first_page_text or "PDW *" in second_page_text:
                                blood_report = file
                                print("This is a blood report.")
                            elif 'SMART REPORT' in first_page_text:
                                smart_report = file
                                print("This is a smart report.")
                            elif 'CERTIFICATE OF VACCINATION' in first_page_text:
                                vaccination_report = file
                                print("This is a vaccination file.")
                            else:
                                others = file
                                print("This is an others file.")


                    except Exception as e:
                        print(f"Error processing file: {file}")
                        print(f"Error details: {str(e)}")

                        # Move the problematic file to the error folder
                        error_folder = pdf_output_dir / "error_pdf"
                        error_folder.mkdir(parents=True, exist_ok=True)
                        # move_to_error_folder(file, error_folder)
                        shutil.copy2(file, error_folder)
                        print(f"This is the problematic file : {file}")

                # Check if at least one file is available for merging
                if xray or optometry or ecg_report or pft or audiometry or vitals or others:
                    merger = PdfMerger()
                    if smart_report:
                        merger.append(smart_report)
                    if vitals:
                        merger.append(vitals)
                    if ecg_report:
                        merger.append(ecg_report)
                    if pft:
                        merger.append(pft)
                    if audiometry:
                        merger.append(audiometry)
                    if optometry:
                        merger.append(optometry)
                    if xray:
                        merger.append(xray)
                    if blood_report:
                        merger.append(blood_report)
                    if vaccination_report:
                        merger.append(vaccination_report)
                    if others:
                        merger.append(others)

                    if len(merger.pages) >= 1:
                        merged_pdf_dir = pdf_output_dir / "7_pages"
                    else:
                        pass

                    base_file_name = (
                        xray.stem.split(".")[0].lower() if xray else
                        optometry.stem.split(".")[0].lower() if optometry else
                        ecg_report.stem.split(".")[0].lower() if ecg_report else
                        pft.stem.split(".")[0].lower() if pft else
                        audiometry.stem.split(".")[0].lower() if audiometry else
                        vitals.stem.split(".")[0].lower() if vitals else
                        vaccination_report.stem.split(".")[0].lower() if vaccination_report else
                        others.stem.split(".")[0].lower() if others else
                        "default_name"
                    )
                    print(base_file_name)

                    merged_file_path = merged_pdf_dir / f"{base_file_name}.pdf"
                    merged_file_path.parent.mkdir(parents=True, exist_ok=True)

                    merger.write(str(merged_file_path))
                    print(f"Merged PDF saved to: {merged_file_path}")

            # Display message box after merging is complete
            total_input_count = len(pdf_files)
            final_pagers = list((pdf_output_dir / "7_pages").glob("*.pdf"))
            total_pdfs = len(final_pagers)

            total_count =  total_pdfs
            # Adding the count logics :
            if total_count <= 3 :
                pdf_files_3_or_less = True
                # updating the count.
                total_pdfs_less_than_3 += 1

            tk.messagebox.showinfo("PDF Merger", f"Total {total_input_count} PDF files merged into {total_count} PDF files successfully!")

            # Display message box with Less than 3 Pages if any
            if pdf_files_3_or_less:
                # file_list = "\n".join(str(file) for file in pdf_files_3_or_less)
                tk.messagebox.showinfo("Missing Files",
                                       f"Total {total_pdfs_less_than_3} merged PDF files have only less then 3 or more then 4 pages.")
            else:
                tk.messagebox.showinfo("No Missing Files", "All merged PDF files have 3 or 4 pages.")
        else:
            tk.messagebox.showwarning("Output Directory", "Output directory not selected.")
    else:
        tk.messagebox.showwarning("Input Directory", "Input directory not selected.")


# def merge_redcliffe_pdf_files():
#     # Prompt user to select input directory
#     input_directory = filedialog.askdirectory(title="Select Input Directory")
#     if not input_directory:
#         tk.messagebox.showwarning("Input Directory", "Input directory not selected.")
#         return

#     # Prompt user to select output directory
#     output_directory = filedialog.askdirectory(title="Select Output Directory")
#     if not output_directory:
#         tk.messagebox.showwarning("Output Directory", "Output directory not selected.")
#         return

#     pdf_dir = Path(input_directory)
#     pdf_output_dir = Path(output_directory)
#     pdf_output_dir.mkdir(parents=True, exist_ok=True)

#     # List all PDF files in the input directory
#     pdf_files = list(pdf_dir.glob("*.pdf"))

#     # Dictionary to categorize files by patient ID
#     patient_files = {}

#     for file in pdf_files:
#         # Extract patient ID from filename
#         parts = file.stem.split("_")
#         patient_id = parts[0].lower()  # Use lowercase for consistency

#         if patient_id not in patient_files:
#             patient_files[patient_id] = {'xray': None, 'optometry': None, 'ecg': None, 'pft': None, 'audiometry': None, 'vitals': None, 'others': None}

#         try:
#             pdf_reader = PdfReader(open(file, "rb"))
#             first_page_text = pdf_reader.pages[0].extract_text() if len(pdf_reader.pages) > 0 else ''
#             second_page_text = pdf_reader.pages[1].extract_text() if len(pdf_reader.pages) > 1 else ''
            
#             print(f"Processing file: {file}")
#             print(f"First Page Text: {first_page_text[:100]}")  # Print first 100 chars for debugging
#             print(f"Second Page Text: {second_page_text[:100]}")  # Print first 100 chars for debugging

#             if "Study Date" in first_page_text and "Report Date" in first_page_text:
#                 patient_files[patient_id]['xray'] = file
#             elif "OPTOMETRY" in first_page_text:
#                 patient_files[patient_id]['optometry'] = file
#             elif "ECG" in second_page_text:
#                 patient_files[patient_id]['ecg'] = file
#             elif "RECORDERS & MEDICARE SYSTEMS" in first_page_text:
#                 patient_files[patient_id]['pft'] = file
#             elif "VITALS" in first_page_text:
#                 patient_files[patient_id]['vitals'] = file
#             elif 'left ear' in first_page_text:
#                 patient_files[patient_id]['audiometry'] = file
#             else:
#                 patient_files[patient_id]['others'] = file

#         except Exception as e:
#             print(f"Error processing file: {file}")
#             print(f"Error details: {str(e)}")
#             # Move the problematic file to the error folder
#             error_folder = pdf_output_dir / "error_pdf"
#             error_folder.mkdir(parents=True, exist_ok=True)
#             move_to_error_folder(file, error_folder)

#     # Process each patient and merge PDFs
#     total_input_count = len(pdf_files)
#     merged_pdfs_count = 0

#     for patient_id, files in patient_files.items():
#         print(f"Merging files for patient ID: {patient_id}")
#         merger = PdfMerger()
#         for category in ['vitals', 'ecg', 'xray', 'pft', 'audiometry', 'optometry', 'others']:
#             if files[category]:
#                 print(f"Appending file: {files[category]}")
#                 merger.append(files[category])
        
#         if len(merger.pages) > 0:
#             merged_pdf_dir = pdf_output_dir / "7_pages"
#             merged_pdf_dir.mkdir(parents=True, exist_ok=True)
            
#             base_file_name = patient_id
#             merged_file_path = merged_pdf_dir / f"{base_file_name}.pdf"
#             merger.write(str(merged_file_path))
#             merger.close()
#             merged_pdfs_count += 1
#             print(f"Merged PDF saved to: {merged_file_path}")

#     # Display message box after merging is complete
#     tk.messagebox.showinfo("PDF Merger", f"Total {total_input_count} PDF files processed. {merged_pdfs_count} PDFs merged successfully!")

# def move_to_error_folder(file, error_folder):
#     try:
#         shutil.move(str(file), error_folder / file.name)
#     except Exception as e:
#         print(f"Error moving file to error folder: {file}")
#         print(f"Error details: {str(e)}")

# if __name__ == "__main__":
#     merge_redcliffe_pdf_files()

def merge_all():
    # Prompt user to select input directory
    input_directory = filedialog.askdirectory(title="Select Input Directory")

    if input_directory:
        # Prompt user to select output directory
        output_directory = filedialog.askdirectory(title="Select Output Directory")

        if output_directory:
            pdf_dir = Path(input_directory)
            pdf_output_dir = Path(output_directory)
            pdf_output_dir.mkdir(parents=True, exist_ok=True)

            # List all PDF files in the input directory
            pdf_files = list(pdf_dir.glob("*.pdf"))
            if pdf_files:
                # Create a PdfFileMerger object
                merger = PdfMerger()
                for pdf_file in pdf_files:
                    merger.append(pdf_file)
                output_file_path = pdf_output_dir / "merged_file.pdf"

                # Write the merged PDF to the output file
                with open(output_file_path, "wb") as output_file:
                    merger.write(output_file)

                print(f"Merged PDF saved to: {output_file_path}")
                tk.messagebox.showinfo("PDF Merger",f"Total {len(pdf_files)} PDF files merged into one PDF successfully!")
            else:
                tk.messagebox.showinfo("No PDF Files", "No PDF files found in the input directory.")
        else:
            tk.messagebox.showwarning("Output Directory", "Output directory not selected.")
    else:
        tk.messagebox.showwarning("Input Directory", "Input directory not selected.")

def rename_pdf_files():
    input_directory = filedialog.askdirectory(title="Select Input Directory")

    if input_directory:
        # Prompt user to select output directory
        output_directory = filedialog.askdirectory(title="Select Output Directory")

        if output_directory:
            input_dir = Path(input_directory)
            output_dir = Path(output_directory)
            output_dir.mkdir(parents=True, exist_ok=True)

            error_dir = output_dir / "error_files"
            error_dir.mkdir(parents=True, exist_ok=True)

            # List all PDF files in the input directory
            pdf_files = list(input_dir.glob("*.pdf"))

            if pdf_files:
                renamed_count = 0
                error_count = 0
                patient_id = ''
                patient_name = ''

                for pdf_file in pdf_files:
                    with open(pdf_file, 'rb') as file:
                        pdf_reader = PyPDF2.PdfReader(file)
                        if len(pdf_reader.pages) > 0:
                            first_page_text = ''
                            second_page_text = ''
                            if len(pdf_reader.pages) == 1:
                                first_page = pdf_reader.pages[0]
                                first_page_text = first_page.extract_text()
                            elif len(pdf_reader.pages) >= 4:
                                first_page = pdf_reader.pages[0]
                                first_page_text = first_page.extract_text()
                            else:
                                second_page = pdf_reader.pages[1]
                                second_page_text = second_page.extract_text()
                                first_page_text = second_page_text

                            print("Here is the extracted text from the page :")
                            print(first_page_text)
                            print("End of the extracted text")
                            print("This is the second page text if exists , ", second_page_text)

                            try:
                                # Condition for renaming the blood report - Himanshu.
                                if "RBC Count" in first_page_text:
                                    complete_patient_name = str(first_page_text).split("Patient Name : ")[1].split("DOB/")[0].strip()
                                    patient_id = complete_patient_name.rsplit(" ",1)[1]
                                    patient_name = complete_patient_name.rsplit(" ",1)[0].split(" ",1)[1].lower()
                                    print("This is the complete Patient Name extracted : ", complete_patient_name)
                                    print("This is the Patient Id : ", patient_id)
                                    print("This is the extracted Patient Name: ", patient_name)
                                # X-RAY
                                elif "Study Date" and "Report Date" in first_page_text:
                                    patient_id = str(first_page_text).split("Patient ID")[1].split(" ")[1].lower().strip()
                                    patient = str(first_page_text).split("Name")[1].split("Date")[0].split(" ")[0].strip().lower()
                                    if "patient" in patient:
                                        patient_name = patient.split("patient")[0].strip()
                                    else:
                                        patient_name = patient
                                    print(patient_id, patient_name)
                                # PFT
                                elif "RECORDERS & MEDICARE SYSTEMS" in first_page_text:
                                    patient_id = str(first_page_text).split("ID     :")[1].split("Age")[0].strip().lower()
                                    patient_name = str(first_page_text).split("Patient: ")[1].split("Refd.By:")[0].split("\n")[0].lower()
                                    if " " in patient_name:
                                        patient_name = patient_name.split(" ")[0]
                                    else:
                                        patient_name = patient_name

                                # ECG GRAPH
                                elif "Acquired on:" in first_page_text:
                                    if "Id :" in first_page_text:
                                        patient_id = str(first_page_text).split("Id :")[1].split(" ")[1].split("\n")[0].strip().lower()
                                        if patient_id == '':
                                            patient_id = str(first_page_text).split("Comments")[1].split("HR")[0].strip()
                                            print("comments", patient_id)

                                    elif "Id:" in first_page_text:
                                        patient_id = str(first_page_text).split("Id:")[1].split(" ")[1].split("\n")[0].strip().lower()
                                        if patient_id == '':
                                            patient_id = str(first_page_text).split("Comments")[1].split("HR")[0].strip()



                                    if "Name :" in first_page_text:
                                        patient_name = str(first_page_text).split("Name :")[1].split("Age")[0].split(" ")[1].strip().lower()
                                    elif "Name:" in first_page_text:
                                        patient_name = str(first_page_text).split("Name:")[1].split("Age")[0].split(" ")[1].strip().lower()
                                    else:
                                        patient_name = 'invalid'

                                elif "ECG" in second_page_text:
                                    patient_id = str(second_page_text).split("Patient ID:")[1].split("Age:")[0].strip().lower()
                                    patient_name = str(second_page_text).split("Name:")[1].split("Patient ID:")[0].strip().lower()
                                elif "left ear" in first_page_text:
                                    patient_id = str(first_page_text).split('Patient ID')[1].split('Age')[0].strip().lower()
                                    patient_name = str(first_page_text).split('Name')[1].split('Patient ID')[0].strip().lower()
                                elif "OPTOMETRY" in first_page_text:
                                    patient_id = str(first_page_text).split("Patient ID:")[1].split("Age:")[0].strip().lower()
                                    patient_name = str(first_page_text).split("Name:")[1].split("Patient ID:")[0].strip().lower()
                                elif "VITALS" in first_page_text:
                                    patient_id = str(first_page_text).split("Patient ID:")[1].split("Age:")[0].strip().lower()
                                    patient_name = str(first_page_text).split("Name:")[1].split("Patient ID:")[0].strip().lower()  


                                renamed_count += 1
                                new_filename = patient_id + "_" + patient_name
                                new_file_path = output_dir / (new_filename + pdf_file.suffix)
                                shutil.copy2(pdf_file, new_file_path)
                                print(f"File renamed and saved: {pdf_file} -> {new_file_path}")

                            except Exception as e:
                                error_count += 1
                                error_file_path = error_dir / pdf_file.name
                                shutil.copy2(pdf_file, error_file_path)
                                print(f"Error processing file {pdf_file}: {str(e)}")

                messagebox.showinfo("Renaming Complete", f"{renamed_count} PDF files have been renamed.")
                if error_count > 0:
                    messagebox.showwarning("Error Files", f"{error_count} PDF files encountered errors. They are saved in the 'error_files' folder.")
            else:
                messagebox.showwarning("No PDF Files", "No PDF files found in the input directory.")
        else:
            messagebox.showwarning("Output Directory", "Output directory not selected.")
    else:
        messagebox.showwarning("Input Directory", "Input directory not selected.")

def remove_illegal_characters(value):
    if isinstance(value, str):
        # Remove characters that are not printable or are control characters
        value = re.sub(r'[\x00-\x1F\x7F-\x9F]', '', value)
    return value

def extract_patient_data():
    input_directory = filedialog.askdirectory(title="Select Input Directory")

    if input_directory:
        output_directory = filedialog.askdirectory(title="Select Output Directory")

        if output_directory:
            input_dir = Path(input_directory)
            output_dir = Path(output_directory)
            output_dir.mkdir(parents=True, exist_ok=True)

            error_dir = output_dir / "error_files"
            error_dir.mkdir(parents=True, exist_ok=True)

            pdf_files = list(input_dir.glob("*.pdf"))
            error_count = 0

            patient_data_ecg = []
            patient_data_ecg1 = []
            patient_data_pft = []
            patient_data_xray = []
            patient_data_optometry = []
            patient_data_vitals = []

            total_ecg_files = 0
            total_ecg_files1 = 0
            total_pft_files = 0
            total_xray_files = 0
            total_optometry_files = 0
            total_vitals_files = 0

            excel_file_path_ecg = ""
            excel_file_path_ecg1 = ""
            excel_file_path_pft = ""
            excel_file_path_xray = ""

            workbook_xray = Workbook()
            sheet_xray = workbook_xray.active
            row_xray = 2

            for pdf_file in pdf_files:
                with open(pdf_file, 'rb') as file:
                    pdf_reader = PyPDF2.PdfReader(file)
                    if len(pdf_reader.pages) > 0:
                        first_page_text = ''
                        second_page_text = ''
                        if len(pdf_reader.pages) == 1:
                            first_page = pdf_reader.pages[0]
                            first_page_text = first_page.extract_text()
                        elif len(pdf_reader.pages) == 2:
                            first_page = pdf_reader.pages[1]
                            first_page_text = first_page.extract_text()
                        else:
                            second_page = pdf_reader.pages[1]
                            second_page_text = second_page.extract_text()
                            first_page_text = second_page_text

                        try:
                            # xray
                            if "Study Date" in first_page_text or "Report Date" in first_page_text:
                                patient_id = str(first_page_text).split("Patient ID")[1].split(" ")[1].lower().strip()
                                patient = str(first_page_text).split("Name")[1].split("Date")[0].split(" ")[0].strip().lower()
                                if "patient" in patient:
                                    patient_name = patient.split("patient")[0].strip()
                                else:
                                    patient_name = patient
                                gender = str(first_page_text).split("Sex")[1].split("Study Date")[0].strip().lower()
                                if 'Yr' or 'Y' or 'yrs' in first_page_text:
                                    if 'Yr' in first_page_text:
                                        age_data = str(first_page_text).split("Age")[1].split("Yr")[0].strip()
                                        if "Days" in age_data:
                                            age = age_data.split("Days")[0]
                                        else:
                                            age = age_data
                                    if 'Y' in first_page_text:
                                        age_data = str(first_page_text).split("Age")[1].split('Y')[0].strip()
                                        if "Days" in age_data:
                                            age = age_data.split("Days")[0]
                                        else:
                                            age = age_data
                                    if 'yrs' in first_page_text:
                                        age_data = str(first_page_text).split("Age")[1].split('yrs')[0].strip()
                                        if "Days" in age_data:
                                            age = age_data.split("Days")[0]
                                        else:
                                            age = age_data

                                test_date = str(first_page_text).split("Study Date")[1].split("\n")[1].split("Time")[1]
                                report_date = str(first_page_text).split("Report Date")[1].split("\n")[1].split("Time")[1]

                                if "Adv: Clinical correlation." not in first_page_text:
                                    findings_data = str(first_page_text).split("IMPRESSION")[1].split("Correlate clinically")[0].split(":")[1].strip()
                                    if "Please" in findings_data:
                                        findings_with_dot = findings_data.split("Please")[0]
                                        if "•" in findings_with_dot:
                                            findings = findings_with_dot.split("•")[1].split(".")[0]
                                        else:
                                            findings = findings_with_dot.split(".")[0]
                                    else:
                                        findings_with_dot = findings_data
                                        if "•" in findings_with_dot:
                                            findings = findings_with_dot.split("•")[1].split(".")[0]
                                        else:
                                            findings = findings_with_dot.split(".")[0]


                                if "Adv: Clinical correlation." in first_page_text:
                                    findings_data1 = str(first_page_text).split("Impression")[1]
                                    if findings_data1:
                                        findings = findings_data1.split("Adv: Clinical correlation.")[0].split(':')[1].strip()


                                if  findings == 'No significant abnormality noted' or findings == 'No significant abnormality':
                                    findings = 'No significant abnormality seen'
                                patient_data_xray.append((patient_id, patient_name, age, gender, test_date, report_date, remove_illegal_characters(findings)))
                                print(patient_id, patient_name, age, gender, test_date, report_date, findings)
                                total_xray_files += 1

                            # Extract ECG data
                            elif "Acquired on:" in first_page_text:
                                patient_id = str(first_page_text).split("Id :")[1].split(" ")[1].split("\n")[0]
                                patient_name = str(first_page_text).split("Name :")[1].split("Age :")[0]
                                patient_age = str(first_page_text).split("Age :")[1].split(" ")[1].split("\n")[0]
                                patient_gender = str(first_page_text).split("Gender :")[1].split("|")[0].strip()
                                heart_rate = str(first_page_text).split("HR:")[1].split("R(II):")[0].strip()
                                report_time = str(first_page_text).split("Acquired on:")[1][12:17]
                                report_date = str(first_page_text).split("Acquired on:")[1][1:11]
                                R_II = str(first_page_text).split("R(II):")[1].split("RR")[0].strip()
                                RR = str(first_page_text).split("RR:")[1].split("PR:")[0].strip()
                                PR = str(first_page_text).split("PR:")[1].split("QRS:")[0].strip()
                                QRS = str(first_page_text).split("QRS:")[1].split("QT:")[0].strip()
                                QT = str(first_page_text).split("QT:")[1].split("QTc:")[0].strip()
                                QTc = str(first_page_text).split("QTc:")[1].split("QT/QTc:")[0].split("QT/")[0].strip()
                                QT_QTc = str(first_page_text).split("QT/QTc:")[1].strip()

                                print(patient_id, patient_name, patient_age, patient_gender, heart_rate, report_time, report_date)
                                patient_data_ecg.append((patient_id, patient_name, patient_age, patient_gender, heart_rate,report_time, report_date, R_II
                                                         ,RR, PR, QRS, QT, QTc, QT_QTc))
                                total_ecg_files += 1

                            #pft
                            elif "RECORDERS & MEDICARE SYSTEMS" in first_page_text:
                                patient_id = str(first_page_text).split("ID")[1].split("Age")[0].split(":")[1].strip()
                                patient_name = str(first_page_text).split("Patient")[1].split("Refd.By:")[0].split(":")[1].strip()
                                patient_age = str(first_page_text).split("Age    :")[1].split("Yrs")[0].strip()
                                gender = str(first_page_text).split("Gender")[1].split("Smoker")[0].split(":")[1].strip()
                                height = str(first_page_text).split("Height :")[1].split("Weight")[0].strip()
                                weight = str(first_page_text).split("Weight")[1].split("Gender")[0].split(":")[1].split("Kgs")[0].strip()
                                date = str(first_page_text).split("Date")[1][1:21].split(":")[1]
                                observation = str(first_page_text).split("Pre Test COPD Severity")[1]
                                patient_data_pft.append((patient_id, patient_name, patient_age, gender, height, weight, date, remove_illegal_characters(observation)))
                                print(patient_id, patient_name, patient_age, gender, height, weight, date, observation)
                                total_pft_files += 1

                            #ECG-REPORTINGBOT
                            elif "ECG" in first_page_text:
                                patient_id = str(first_page_text).split('Patient ID:')[1].split('Age:')[0].strip()
                                patient_name = str(first_page_text).split("Name:")[1].split("Patient ID:")[0].strip()
                                age = str(first_page_text).split("Age:")[1].split('Gender:')[0].strip()
                                gender = str(first_page_text).split("Gender:")[1].split("Test date:")[0].strip()
                                test_date = str(first_page_text).split("Test date:")[1].split('Report date:')[0].strip()
                                report_date = str(first_page_text).split("Report date:")[1].split('ECG')[0].strip()
                                heart_rate = str(first_page_text).split("Heart rate is")[1].split("BPM.")[0].strip()
                                findings = str(first_page_text).split("2.")[1].split('.')[0].strip()
                                patient_data_ecg1.append((patient_id, patient_name, age, gender, test_date, report_date, heart_rate, remove_illegal_characters(findings)))
                                total_ecg_files1 += 1


                            #XRAY-REPORTINGBOT
                            else:
                                patient_id = str(first_page_text).split('Patient ID:')[1].split('Age:')[0].strip()
                                patient_name= str(first_page_text).split('Name:')[1].split('Patient ID:')[0].strip()
                                age = str(first_page_text).split('Age:')[1].split('Gender:')[0].strip()
                                gender = str(first_page_text).split('Gender:')[1].split('Test date:')[0].strip()
                                test_date = str(first_page_text).split('Test date:')[1].split('Report date:')[0].strip()
                                report_date = str(first_page_text).split('Report date:')[1].split('X-RAY')[0].strip()
                                findings_data = str(first_page_text).split('IMPRESSION:')[1].split("Dr.")[0]
                                print(findings_data)

                                if "•" in findings_data:
                                    findings = findings_data.split("•")[1].split(".")[0].strip()
                                else:
                                    findings = findings_data
                                print(patient_id, patient_name, age, gender, report_date, remove_illegal_characters(findings_data))
                                patient_data_xray.append((patient_id, patient_name, age, gender, test_date, report_date, remove_illegal_characters(findings)))
                                total_xray_files += 1


                        except IndexError as e:
                            error_count += 1
                            error_file_path = error_dir / pdf_file.name
                            shutil.copy2(pdf_file, error_file_path)
                            print(f"Error processing file {pdf_file}: Invalid PDF Format")

            if total_ecg_files > 0:
                workbook_ecg = openpyxl.Workbook()
                sheet_ecg = workbook_ecg.active

                sheet_ecg['A1'] = 'patient_id'
                sheet_ecg['B1'] = 'name'
                sheet_ecg['C1'] = 'age'
                sheet_ecg['D1'] = 'gender'
                sheet_ecg['E1'] = 'heart_rate'
                sheet_ecg['F1'] = 'report_time'
                sheet_ecg['G1'] = 'report_date'
                sheet_ecg['H1'] = 'R_II'
                sheet_ecg['I1'] = 'RR'
                sheet_ecg['J1'] = 'PR'
                sheet_ecg['K1'] = 'QRS'
                sheet_ecg['L1'] = 'QT'
                sheet_ecg['M1'] = 'QTc'
                sheet_ecg['N1'] = 'QT_QTc'

                for row, data in enumerate(patient_data_ecg, start=2):
                    sheet_ecg.append(data)

                excel_file_path_ecg = os.path.join(output_dir, "patient_data_ecg.xlsx")
                workbook_ecg.save(excel_file_path_ecg)

            if total_ecg_files1 > 0:
                workbook_ecg = openpyxl.Workbook()
                sheet_ecg = workbook_ecg.active

                sheet_ecg['A1'] = 'patient_id'
                sheet_ecg['B1'] = 'name'
                sheet_ecg['C1'] = 'age'
                sheet_ecg['D1'] = 'gender'
                sheet_ecg['E1'] = 'test_date'
                sheet_ecg['F1'] = 'report_date'
                sheet_ecg['G1'] = 'heart_rate'
                sheet_ecg['H1'] = 'findings'

                for row, data in enumerate(patient_data_ecg1, start=2):
                    sheet_ecg.append(data)

                excel_file_path_ecg = os.path.join(output_dir, "patient_data_ecg.xlsx")
                workbook_ecg.save(excel_file_path_ecg)


            if total_pft_files > 0:
                workbook_pft = openpyxl.Workbook()
                sheet_pft = workbook_pft.active

                sheet_pft['A1'] = 'patient_id'
                sheet_pft['B1'] = 'name'
                sheet_pft['C1'] = 'age'
                sheet_pft['D1'] = 'gender'
                sheet_pft['E1'] = 'height'
                sheet_pft['F1'] = 'weight'
                sheet_pft['G1'] = 'date'
                sheet_pft['H1'] = 'observation'

                for row, data in enumerate(patient_data_pft, start=2):
                    sheet_pft.append(data)

                excel_file_path_pft = os.path.join(output_dir, "patient_data_pft.xlsx")
                workbook_pft.save(excel_file_path_pft)

            if total_xray_files > 0:
                workbook_xray = openpyxl.Workbook()
                sheet_xray = workbook_xray.active

                sheet_xray['A1'] = 'patient_id'
                sheet_xray['B1'] = 'name'
                sheet_xray['C1'] = 'age'
                sheet_xray['D1'] = 'gender'
                sheet_xray['E1'] = 'test_date'
                sheet_xray['F1'] = 'report_date'
                sheet_xray['G1'] = 'Findings'


                for row, data in enumerate(patient_data_xray, start=2):
                    sheet_xray.append(data)

                for row in range(2, len(patient_data_xray) + 2):
                    cell = sheet_xray.cell(row=row, column=7)
                    findings = cell.value
                    if "No significant abnormality seen"  in findings:
                        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00", fill_type="solid")  # Green
                    else:
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")  # Red

                excel_file_path_xray = os.path.join(output_dir, "patient_data_xray.xlsx")
                workbook_xray.save(excel_file_path_xray)

            message = f"Total {total_ecg_files1} ECG and {total_pft_files} PFT and {total_xray_files} XRAY data files have been extracted and saved successfully.\n\n"
            message += f"ECG Output File: {excel_file_path_ecg}\n\nPFT Output File: {excel_file_path_pft}\n\nXRAY Output File: {excel_file_path_xray}"
            messagebox.showinfo("Patient Data Extractor", message)

        else:
            messagebox.showwarning("Output Folder Not Selected", "Output folder not selected.")
    else:
        messagebox.showwarning("Input Folder Not Selected", "Input folder not selected.")

def check_pdf_files():
    # Asking for the I/P Directory ( files that need to be checked with the excel.)
    pdf_folder_path = filedialog.askdirectory(title="Select Merged PDF Folder", mustexist=True)
    if not pdf_folder_path:
        print("Merged PDF folder not selected.")
        return

    # Asking for the excel directory (the excel which is used for comparison.)
    excel_file_path = filedialog.askopenfilename(title="Select Excel Sheet", filetypes=[("Excel Files", "*.xlsx;*.xls")])
    if not excel_file_path:
        print("Excel sheet not selected.")
        return

    # Asking for the O/P Directory (here the compared excel will come.)
    output_directory = filedialog.askdirectory(title="Select Output Directory")
    if not output_directory:
        print("Output directory not selected.")
        return

    wb = Workbook()
    ws = wb.active

    # Add headers to the worksheet
    headers = ["patient_id", "patient_name", "age", "gender", "date", "ECG_GRAPH/ECG_REPORT", "XRAY_REPORT",
               "XRAY_IMAGE", "PFT", "AUDIOMETRY", "OPTOMETRY", "VITALS", "PROBLEM"]
    for col_num, header in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_num)
        cell.value = header
        cell.font = Font(bold=True)

    comparison_df = pd.read_excel(excel_file_path)

    for _, excel_row in comparison_df.iterrows():
        pdf_id_prefix = str(excel_row['patient_id']).lower()

        pdf_files = [file.lower() for file in os.listdir(pdf_folder_path) if file.split("_")[0].lower() == pdf_id_prefix]

        # Initialize modality matching list outside the PDF page loop
        modality_match_list = []
        problem_list = []
        if not pdf_files:
            # No matching PDF file found for patient ID
            problem_list.append("Pdf file is missing")
            modality_match_list = ["No"] * 7

            # Write the results to the worksheet
            row_data = [
                           str(excel_row["patient_id"]).lower(),
                           str(excel_row["patient_name"]).split(" ")[0].lower(),
                           str(excel_row.get("age", "")).strip(),
                           str(excel_row["gender"]).strip(),
                           str(excel_row["date"]).strip(),
                       ] + modality_match_list + [', '.join(problem_list)]
            ws.append(row_data)

            current_row = ws.max_row

            yes_columns = [5, 6, 7, 8, 9, 10, 11, 12]  # columns E to L are modality columns

            # Apply fill color to cells based on "Yes" or "No"
            for col_num in range(5, 13):  # Columns E to L
                cell = ws.cell(row=current_row, column=col_num)
                if cell.value == "Yes":
                    cell.fill = PatternFill(start_color="00FF00", end_color="00FF00",
                                            fill_type="solid")  # Green color
                elif cell.value == "No":
                    cell.fill = PatternFill(start_color="FF0000", end_color="FF0000",
                                            fill_type="solid")  # Red color
                    

        # To check whether pdf files are there or not.
        if pdf_files:
            pdf_file = pdf_files[0]
            pdf_path = os.path.join(pdf_folder_path, pdf_file)

            # Extract patient data from the Excel row
            patient_data_excel = {
                "patient_id": str(excel_row["patient_id"]).lower().strip(),
                "patient_name": str(excel_row["patient_name"]).lower().strip(),
                "age": str(excel_row.get("age", "")).strip(),
                "gender": str(excel_row["gender"]).strip().lower(),
                "date": str(excel_row["date"]).split(" ")[0]
            }
            print(patient_data_excel)

            # Main Logic for comparison.
            try:
                for modality in ["ECG_GRAPH/ECG_REPORT", "XRAY_REPORT", "XRAY_IMAGE", "PFT", "AUDIOMETRY","OPTOMETRY", "VITALS"]:
                    modality_match = False
                    patient_id = None
                    patient_name = None
                    age = None
                    gender = None
                    report_date = None
                    # Open the PDF file for the current row
                    try:
                        pdf_reader = PdfReader(open(pdf_path, "rb"))
                    except Exception as e:
                        print(f"Error processing PDF file {pdf_file}: {str(e)}")
                        error_folder_path = os.path.join(output_directory, "error")
                        os.makedirs(error_folder_path, exist_ok=True)
                        shutil.move(pdf_path, os.path.join(error_folder_path, pdf_file))
                        continue
                    # Iterate through the PDF pages
                    for page_num in range(len(pdf_reader.pages)):
                        page = pdf_reader.pages[page_num]
                        page_text = page.extract_text()

                        missing_modalities = []
                        print("this is the start of page text.")
                        print("Page no. ", page_num)
                        print(page_text)
                        print("This is the end of page text.")

                        # Checking the ECG details.
                        print("Checking if ecg is there or not.")
                        try:
                            print("Inside the try block of ecg.")
                            if modality == "ECG_GRAPH/ECG_REPORT" and "ECG" in page_text:
                                print("confirmed that it is a ecg file.")
                                patient_name = str(page_text).split("Name:")[1].split("Patient ID:")[0].strip().lower()
                                # if patient_name.count(" ") == 1:
                                #     patient_name = patient_name.strip().lower()
                                # else:
                                #     patient_name = patient_name.split(" ")[1].lower().strip()

                                patient_id = str(page_text).split("Patient ID:")[1].split("Age")[0].strip().lower()
                                age = str(page_text).split("Age:")[1].split("Gender")[0].strip()
                                gender = str(page_text).split("Gender:")[1].split("Test")[0].strip().lower()
                                report_date = str(page_text).split("Report date:")[1].split("ECG")[0].strip().lower()
                                print("Printing the details of ECG :")
                                print("Patient Id", patient_id)
                                print("Patient Name", patient_name)
                                print("Age", age)
                                print("Gender", gender)
                                print("Report Date", report_date)


                                print("ECG REPORT/ECG GRAPH", patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                        patient_name == patient_data_excel["patient_name"] and
                                        age == patient_data_excel["age"] and
                                        gender == patient_data_excel["gender"] and
                                        report_date == patient_data_excel["date"]):
                                    modality_match = True
                                    break
                                 

                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                        # Checking if X-RAY file is present (for stradus.)
                        # print("Now checking the details of xray files if present.")
                        # try:
                        #     print("Inside the try block of xray.")
                        #     if modality == "XRAY_REPORT" and "Study Date" and "Report Date" in page_text:
                        #         print("It is confirmed that it is a xray file.")
                        #         patient_id = str(page_text).split("Patient ID")[1].split(" ")[1].lower().strip()
                        #         patient = str(page_text).split("Name")[1].split("Date")[0].split(" ")[0].strip().lower()
                        #         if "patient" in patient:
                        #             patient_name = patient.split("patient")[0].strip()
                        #         else:
                        #             patient_name = patient
                        #         age = str(page_text).split("Age")[1].split("Yr")[0].strip()
                        #         gender = str(page_text).split("Sex")[1].split("Study Date")[0].strip().lower()
                        #         date = str(page_text).split("Study Date")[1].split("\n")[1].split("Time")[1].strip()
                        #         input_date = datetime.strptime(date, "%d %b %Y")
                        #         report_date = input_date.strftime("%Y-%m-%d")

                        #         print("These are the extracted data of the xray.")
                        #         print('XRAY', patient_id, patient_name, age, report_date)
                        #         print("This is the date extracted :", date)
                        #         print("This is the i/p date :", input_date)

                        #         # Compare with Excel data
                        #         if (patient_id == patient_data_excel["patient_id"] and
                        #                 patient_name == patient_data_excel["patient_name"] and
                        #                 age == patient_data_excel["age"] and
                        #                 gender == patient_data_excel["gender"] and
                        #                 report_date == patient_data_excel["date"]):
                        #             modality_match = True
                        #             break
                                 
                        # except IndexError as ie:
                        #     print(f"IndexError: {str(ie)}. Skipping page processing.")
                        #     continue

                        # try:
                        #     if modality == "XRAY_IMAGE" and "Page 2 of 2" in page_text:
                        #         if "Page 2 of 2" in page_text:
                        #             modality_match = True
                        #             break
                        # except IndexError as ie:
                        #     print(f"IndexError: {str(ie)}. Skipping page processing.")
                        #     continue

                        # checking for pft.
                        try:
                            print("inside the try block of pft.")
                            if modality == "PFT" and "RECORDERS & MEDICARE SYSTEMS" in page_text:
                                print("it confirms that it is a pft file.")
                                patient_name = str(page_text).split("Patient: ")[1].split("Refd.By:")[0].split("\n")[0].lower()
                                if " " in patient_name:
                                    patient_name = patient_name.split(" ")[0]
                                else:
                                    patient_name = patient_name
                                patient_id = str(page_text).split("ID     :")[1].split("Age")[0].strip().lower()
                                age = str(page_text).split("Age    :")[1].split("Yrs")[0].strip()
                                if "Smoker" in page_text:
                                    gender = str(page_text).split("Gender   :")[1].split("Smoker")[0].strip().lower()
                                else:
                                    gender = str(page_text).split("Gender   :")[1].split("Eth. Corr:")[0].strip().lower()
                                date = str(page_text).split("Date   :")[1][1:13].strip().lower()
                                if len(date) == 10:
                                    input = date
                                else:
                                    input_date = datetime.strptime(date, "%d-%b-%Y")

                                report_date = input_date.strftime("%Y-%m-%d")

                                print('PFT', patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                        patient_name == patient_data_excel["patient_name"] and
                                        age == patient_data_excel["age"] and
                                        gender == patient_data_excel["gender"] and
                                        report_date == patient_data_excel["date"]):
                                    modality_match = True
                                    break
                                 

                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                        # checking for audio.
                        try:
                            print("inside the try block of audiometry.")
                            if modality == "AUDIOMETRY" and "left ear" in page_text:
                                print("it is confirmation that this is a audiometry file.")
                                data = str(page_text)
                                patient_name = str(page_text).split("Name")[1].split("Patient ID")[0].strip().lower()
                                patient_id = str(page_text).split("Patient ID")[1].split("Age")[0].strip().lower()
                                age = str(page_text).split("Age")[1].split("Gender")[0].strip()
                                
                                gender = str(page_text).split("Gender")[1].split("Test")[0].strip().lower()
                                report_date = str(page_text).split('Report date')[1].strip().lower()

                                print('AUDIOMETRY', patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                        patient_name == patient_data_excel["patient_name"] and
                                        age == patient_data_excel["age"] and
                                        gender == patient_data_excel["gender"]):
                                    modality_match = True
                                    break
                                 

                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                        # Checking for opto.
                        try:
                            print("Inside the try block of optometry.")
                            if modality == "OPTOMETRY" and "OPTOMETRY REPORT" in page_text:
                                print("This is confirmed that this is a opto file.")
                                patient_name = str(page_text).split("Name:")[1].split("Age:")[0].strip().lower()
                                patient_id = str(page_text).split("Patient ID:")[1].split("Patient Name:")[0].strip().lower()
                                age = str(page_text).split("Age:")[1].split("Gender")[0].strip()
                                gender = str(page_text).split("Gender:")[1].split("Test")[0].strip().lower()
                                report_date = str(page_text).split("Report Date:")[1].split("OPTOMETRY")[0].strip().lower()

                                print("These are the opto patient details :")
                                print("Patient Id", patient_id)
                                print("Patient Name", patient_name)
                                print("Age", age)
                                print("Gender", gender)
                                print("Report Date", report_date)
                                

                                print('OPTOMETRY', patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                        patient_name == patient_data_excel["patient_name"] and
                                        age == patient_data_excel["age"] and
                                        gender == patient_data_excel["gender"] and
                                        report_date == patient_data_excel["date"]):
                                    modality_match = True
                                    break
                                else:
                                    if (patient_id != patient_data_excel["patient_id"] and
                                            patient_name != patient_data_excel["patient_name"] and
                                            age != patient_data_excel["age"] and
                                            gender != patient_data_excel["gender"] and
                                            report_date != patient_data_excel["date"]):
                                        problem_list.append(f' {modality}: All the data are incorrect')

                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                        # Checking for vitals.
                        try:
                            print("inside the try block of vitals.")
                            if modality == "VITALS" and "VIT" in page_text:
                                print("it confirms that it is a vitals file.")
                                patient_id = str(page_text).split("Patient ID:")[1].split("Patient Name:")[0].strip().lower()
                                patient_name = str(page_text).split("Patient Name:")[1].split("Age")[0].strip().lower()
                                age = str(page_text).split("Age:")[1].split("Gender")[0].strip()
                                gender = str(page_text).split("Gender:")[1].split("Test")[0].strip().lower()
                                report_date = str(page_text).split("Report Date:")[1].split("VITALS")[0].strip().lower()
                                print('VITALS', patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                    patient_name == patient_data_excel["patient_name"] and
                                    age == patient_data_excel["age"] and
                                    gender == patient_data_excel["gender"] and
                                    report_date == patient_data_excel["date"]):

                                    modality_match = True
                                    break
                                else:
                                    if (patient_id != patient_data_excel["patient_id"] and
                                            patient_name != patient_data_excel["patient_name"] and
                                            age != patient_data_excel["age"] and
                                            gender != patient_data_excel["gender"] and
                                            report_date != patient_data_excel["date"]):
                                        problem_list.append(f' {modality}: All the data are incorrect')
                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                        # Checking for X-Ray (Reporting Bot.)
                        try:
                            print("inside the try block of xray.")
                            if modality == "XRAY_REPORT" and "X-RAY" in page_text:
                                print("it confirms that it is a xray file.")
                                patient_id = str(page_text).split("Patient ID:")[1].split("Age:")[0].lower().strip()
                                patient_name = str(page_text).split("Name:")[1].split("Patient ID:")[0].strip().lower()
                                age = str(page_text).split("Age:")[1].split("YGender:")[0].strip()
                                if age.startswith("0"):
                                    age = age.split("0")[1]
                                gender = str(page_text).split("Gender:")[1].split("Test date:")[0].strip().lower()
                                report_date = str(page_text).split("Report date:")[1].split("X-RAY")[0].strip().lower()

                                print('XRAY BOT', patient_id, patient_name, age, gender, report_date)

                                # Compare with Excel data
                                if (patient_id == patient_data_excel["patient_id"] and
                                        patient_name == patient_data_excel["patient_name"] and
                                        age == patient_data_excel["age"] and
                                        gender == patient_data_excel["gender"] and
                                        report_date == patient_data_excel["date"]):
                                    modality_match = True
                                    break
                                else:
                                    if (patient_id != patient_data_excel["patient_id"] and
                                            patient_name != patient_data_excel["patient_name"] and
                                            age != patient_data_excel["age"] and
                                            gender != patient_data_excel["gender"] and
                                            report_date != patient_data_excel["date"]):
                                        problem_list.append(f' {modality}: All the data are incorrect')
                        except IndexError as ie:
                            print(f"IndexError: {str(ie)}. Skipping page processing.")
                            continue

                    issues = []
                    if patient_id == patient_data_excel["patient_id"] or patient_name == patient_data_excel["patient_name"] or age == patient_data_excel["age"] or gender == patient_data_excel["gender"] or report_date == patient_data_excel["date"]:
                        if not modality_match:
                            if patient_id != patient_data_excel["patient_id"]:
                                issues.append("ID")
                            if patient_name != patient_data_excel["patient_name"]:
                                issues.append("Name")
                            if age != patient_data_excel["age"]:
                                issues.append("Age")
                            if gender != patient_data_excel["gender"]:
                                issues.append("Gender")
                            if report_date != patient_data_excel["date"]:
                                issues.append("Date")

                    # Append the modality 2 corresponding issues to the problem_list
                    if issues:
                        problem_list.append(f"{modality}: {', '.join(issues)}")
                    modality_match_list.append("Yes" if modality_match else "No")

                # Write the results to the worksheet
                row_data = [
                    patient_data_excel["patient_id"],
                    patient_data_excel["patient_name"],
                    patient_data_excel["age"],
                    patient_data_excel["gender"],
                    patient_data_excel["date"]
                ] + modality_match_list  + [', '.join(problem_list)]
                ws.append(row_data)

                current_row = ws.max_row

                # Define the column indices for "Yes" and "No" values
                yes_columns = [5, 6, 9, 10, 11, 12]  # Assuming columns F to M are modality columns

                # Apply fill color to cells based on "Yes" or "No"
                for col_num in range(5, 13):  # Columns E to M
                    cell = ws.cell(row=current_row, column=col_num)
                    if cell.value == "Yes":
                        cell.fill = PatternFill(start_color="00FF00", end_color="00FF00",
                                                fill_type="solid")  # Green color
                    elif cell.value == "No":
                        cell.fill = PatternFill(start_color="FF0000", end_color="FF0000",fill_type="solid")

            except Exception as e:
                print(f"Error processing PDF file {pdf_file}: {str(e)}")
                error_folder_path = os.path.join(output_directory, "error")
                os.makedirs(error_folder_path, exist_ok=True)
                shutil.move(pdf_path, os.path.join(error_folder_path, pdf_file))
                continue
        else:
            print(f"No matching PDF file found for patient ID: {pdf_id_prefix}")

    # Save the workbook to the output directory
    output_filename = "patient_data_comparison.xlsx"
    wb.save(os.path.join(output_directory, output_filename))
    print("Data comparison completed.")
    messagebox.showinfo("Process Completed", "Data comparison completed.")

def sanitize_filename(filename):
    # Remove invalid characters from the filename
    return re.sub(r'[\\/:*?"<>|]', '_', filename)

def split_patient_file():
    input_directory = filedialog.askdirectory(title="Select Input Directory")

    if input_directory:
        output_directory = filedialog.askdirectory(title="Select Output Directory")

        if output_directory:
            pdf_dir = Path(input_directory)
            pdf_output_dir = Path(output_directory)
            pdf_output_dir.mkdir(parents=True, exist_ok=True)

            pdf_files = list(pdf_dir.glob("*.pdf"))

            if pdf_files:
                for input_pdf_path in pdf_files:
                    try:
                        # Open the merged PDF file
                        with open(input_pdf_path, 'rb') as pdf_file:
                            pdf_reader = PyPDF2.PdfReader(pdf_file)

                            # Create a subdirectory for the patient if it doesn't exist
                            patient_id = sanitize_filename(input_pdf_path.stem)
                            patient_dir = pdf_output_dir / patient_id
                            patient_dir.mkdir(parents=True, exist_ok=True)

                            # Loop through each page in the PDF and save them as individual PDF files
                            for page_number in range(len(pdf_reader.pages)):
                                pdf_writer = PyPDF2.PdfWriter()
                                pdf_writer.add_page(pdf_reader.pages[page_number])

                                # Extract text data from the page to determine the modality
                                page_text = pdf_reader.pages[page_number].extract_text()
                                modality = None
                                images = []
                                is_ecg = False
                                count = 0

                                # Adding logs so that i can identify the page. - himanshu.
                                print("Page text is below")
                                print(page_text)
                                print("End of page text.")

                                if "SMART REPORT" in page_text:
                                    print("This is a smart report.")
                                    modality = "Smart_Report"
                                elif "X-RAY" in page_text:
                                    print("This is a xray file.")
                                    modality = 'Xray_Report'
                                elif "RECORDERS & MEDICARE SYSTEMS" in page_text:
                                    print("This is a pft file.")
                                    modality = 'PFT'
                                elif "Page 2 of 2" in page_text:
                                    print("This is an xray image.")
                                    modality = 'Xray_Image'
                                elif "OPTOMETRY" in page_text:
                                    print("This is a optometry file.")
                                    modality = 'Optometry'
                                elif "left ear" in page_text:
                                    print("This is a audiometry file.")
                                    modality = 'Audiometry'
                                elif "ECG" in page_text:
                                    print("This is an ECG file.")
                                    modality = 'ECG'
                                elif is_ecg == True and count == 1:
                                    print("This is another others image.")
                                    modality = 'Other2'
                                    count += 1
                                elif is_ecg == True and count == 2:
                                    print("This is the 3rd others file.")
                                    modality = 'Others3'
                                elif page_text == '':
                                    print("This is a others image.")
                                    modality = 'Others'
                                    is_ecg = True
                                    count += 1
                                elif "VITALS" in page_text:
                                    print("This is a vitals file.")
                                    modality = 'Vitals'
                                elif "RBC Count" in page_text:
                                    print("This is a blood Report.")
                                    modality = 'BloodReport'
                                else:
                                    print("This is a dr. consultation  file.")
                                    modality = 'Dr.Consultation'

                                # I have to add a proper logic to separate the images separately.

                                if modality:
                                    output_file_path = patient_dir / f'{patient_id}_{modality}.pdf'
                                else:
                                    output_file_path = patient_dir / f'{patient_id}_page_{page_number + 1}.pdf'

                                # Save the individual page as a PDF file with the new name
                                with open(output_file_path, 'wb') as output_file:
                                    pdf_writer.write(output_file)

                            print(f"PDF files for patient {patient_id} split and renamed successfully.")
                    except Exception as e:
                        print(f"Error processing {input_pdf_path}: {str(e)}")
                        continue  # Skip this file and continue with the next

                print("PDF files processed.")
            else:
                print("No PDF files found in the input directory.")

        else:
            print("Output directory not selected.")
    else:
        print("Input directory not selected.")


def check_ecg_files():
    # o/p = "Here i will write logic to check the ecg files."
    print(f"Here i will write the logic to check the ecg files.")

# def dcm_to_pdf_converter():
#     input_directory = filedialog.askdirectory(title="Select Input Directory")
#
#     if input_directory:
#         # Prompt user to select output directory
#         output_directory = filedialog.askdirectory(title="Select Output Directory")
#
#         if output_directory:
#             input_dir = Path(input_directory)
#             output_dir = Path(output_directory)
#             output_dir.mkdir(parents=True, exist_ok=True)
#
#             error_dir = output_dir / "error_files"
#             error_dir.mkdir(parents=True, exist_ok=True)
#
#             pdf_files = list(input_dir.glob("*.pdf"))
#             error_count = 0
#
#
#
#         else:
#             messagebox.showwarning("Output Folder Not Selected", "Output folder not selected.")
#     else:
#         messagebox.showwarning("Input Folder Not Selected", "Input folder not selected.")

# Create the main window
window = tk.Tk()
window.title("Camp - Automation Tools")
# Set the window dimensions and position it on the screen
window.geometry("1000x500+200-100")


redcliffe_label = tk.Label(window, text="Merge Pdf Files", font=("Arial", 16, "bold"))
redcliffe_label.place(x=580, y=10, anchor='ne')

# Adding the label of Merge All files button .
merge_all_files = tk.Label(window, text="Merge All PDF Files", font=("Arial", 16, "bold"))
merge_all_files.place(x=600, y=130, anchor='ne')

merge_redcliffe_button1 = tk.Button(window, bg='blue', fg='white', activebackground='darkblue', activeforeground='white', padx=30, pady=10, relief='raised', text="Merge PDF Files", command=merge_redcliffe_pdf_files, font=("Arial", 12, "bold"))
merge_redcliffe_button2 = tk.Button(window, bg='magenta', fg='black', activebackground='gold', activeforeground='black', padx=30, pady=10, relief='raised', text="Merge All PDF Files", command=merge_all, font=("Arial", 12, "bold"))
merge_redcliffe_button1.place(x=600, y=58, anchor='ne')
merge_redcliffe_button2.place(x=623, y=178, anchor='ne')

pdf_rename_label = tk.Label(window, text="File Renaming System", font=("Arial", 16, "bold"))
pdf_rename_label.pack(pady=10, padx=20, anchor='w')

pdf_rename_button1 = tk.Button(window, bg='orange', fg='black', activebackground='darkblue', activeforeground='white', padx=30, pady=10, relief='raised', text="Rename PDF Files", command=rename_pdf_files, font=("Arial", 12, "bold"))
pdf_rename_button1.pack(pady=8, padx=20, anchor='w')

generate_excel_label = tk.Label(window, text="Data Extracting System", font=("Arial", 16, "bold"))
generate_excel_label.place(x=259, y=130, anchor='ne')

generate_excel_button = tk.Button(window, bg='pink',fg='black', activebackground='darkblue', activeforeground='white',padx=30, pady=10, relief='raised', text="Generate Patient Excel", command=extract_patient_data, font=("Arial", 12, "bold"))
generate_excel_button.place(x=262, y=180, anchor='ne')

check_pdf_File = tk.Label(window, text="Check Pdf Files", font=("Arial", 16, "bold"))
check_pdf_File.place(x=920, y=10, anchor='ne')

check_pdf_button = tk.Button(window, bg='green',fg='black', activebackground='darkblue', activeforeground='white',padx=30, pady=10, relief='raised', text="Check Pdf Files", command=check_pdf_files, font=("Arial", 12, "bold"))
check_pdf_button.place(x=956, y=57, anchor='ne')

check_pdf_File = tk.Label(window, text="Split Pdf Files", font=("Arial", 16, "bold"))
check_pdf_File.place(x=903, y=130, anchor='ne')

check_pdf_button = tk.Button(window, bg='yellow',fg='black', activebackground='darkblue', activeforeground='white',padx=30, pady=10, relief='raised', text="Split Pdf Files", command=split_patient_file, font=("Arial", 12, "bold"))
check_pdf_button.place(x=940, y=175, anchor='ne')

check_ecg_files_label = tk.Label(window, text="Check ECG Files", font=("Arial", 16, "bold"))
check_ecg_files_label.place(x = 580, y=250, anchor='ne')

check_ecg_files_button = tk.Button(window, bg='red',fg='black', activebackground='darkred', activeforeground='white',padx=30, pady=10, relief='raised', text="Check ECG Files", command=check_ecg_files, font=("Arial", 12, "bold"))
check_ecg_files_button.place(x=605, y=310, anchor='ne')


# dcm_to_pdf = tk.Label(window, text="Reports Observation", font=("Arial", 16, "bold"))
# dcm_to_pdf.place(x=233, y=255, anchor='ne')
#
# dcm_to_pdf_button = tk.Button(window, bg='red',fg='black', activebackground='darkblue', activeforeground='white',padx=30, pady=10, relief='raised', text="GET REPORTS OBSERVATION", command=dcm_to_pdf_converter, font=("Arial", 12, "bold"))
# dcm_to_pdf_button.place(x=328, y=300, anchor='ne')



window.mainloop()






