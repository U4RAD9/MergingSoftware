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
-- Himanshu Jangid.