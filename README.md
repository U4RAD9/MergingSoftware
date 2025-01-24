* To run the merging software i have to use : py merge_pdfs.py
* To create the .exe i have to simply run : pyinstaller merge_pdfs.spec

This is the Updated Merging Software. 
I've fixed the following things as of 23rd September 2024 :
1. Fixed the renaming of Blood Reports, ECG Reports , and PFT too.
2. Fixed the merging of all the files of a particular patient including the Dr. consultation file and the Smart Reporting File.
3. Fixed the splitting logic , added both the dr.consultation file and the smart report.
4. Now , I have to add the logic which checks the formatting issue in ecg files and then they upload it to bot.

18 Oct 24:
1. I faced an issue where the key was not matching in the individual patient report merging logic,so i've matched it by converting in lowercase.

12 Nov 24:
1. Operations team needed a new requirement, for adding the vaccination report too, so added that code in the original code.

21 Nov 24 :
1. There was one issue in the check pdf files functionality , the age extracted from the excel from which we need to compare was coming as float, so fixed it to get it as integer and then pass it as string.s

10 Jan 25 :
1. I have changed the logic of merging a person's individual file after version 4 ( i.e. merge_redcliffe_pdf_files fxn), where i have made some specific changes to handle the naming conflicts and the exceptions.
2. Added the logic to give a message displaying the error in any pdf file name format.
3. Also , Adding each and every others file that belongs to a particular id.
4. Fixed the blood report issue (previously it was getting added using "PDW *" ,now also added "PDW").
5. Also, commented a different optimized coding logic, which needs time to continue working on it.

22 Jan 25:
1. I've created another function which takes input as all the merged files which means all files for a particular patient, where it first tells that which files are present for a particular patient, and than gives the data present in the files in the form of excel, named as generate_excel_for_merged_files.
2. The respective button for above function is 'Generate Excel For Merged Files'.
3. I've also changed the function which creates excel for the patients where as of now only xray, ecg , and pft excels were coming as output. Now, that will also have Audio, Opto, Vitals, Blood and Other Files(mostly dr. consultation).
-- Himanshu Jangid.