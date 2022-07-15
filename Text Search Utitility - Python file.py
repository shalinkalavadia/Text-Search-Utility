#!/usr/bin/env python
# coding: utf-8

# In[1]:


"""  FINAL WITH TASK 1  """

"""  THIS FINAL PROGRAM WILL NOT WORK WITH .DOC FILES , BUT WILL WORK WITH SCANNED PDFs """

# VERSION - 3 :- PRIVILEDGES EQUATION ALSO INCLUDED

import subprocess as sp                # Find Hostname from CMD from computer
import pandas as pd                    # For dataframes
import os                              # To collect all the files from OS
import cv2                             # Text extraction from IMAGE - 1 library
import pytesseract                     # Text extraction from IMAGE - 2 library
import fitz                            # Extract text from PDF file
import msoffcrypto                     # Open the password protected excel file by inserting password. - 1
import io                              # Open the password protected excel file by inserting password. - 2
import docx2txt                        # Extract text from WORD file 
import shutil                          # delete the folder containing files and data
from PyPDF2 import PdfFileMerger
from pdf2image import convert_from_path
from PIL import Image

##################################################
# Here, deleting all the previously generated text files(dir1), then all extra files (dir2) and lastly final output(dir3) 

dir1 = "D:\Innovation Team\Text Search Utility\Output\Converted Text Files"
for f1 in os.listdir(dir1):
    os.remove(os.path.join(dir1, f1))

dir2 = "D:\Innovation Team\Text Search Utility\Output\Converted Searchable PDF"
for f2 in os.listdir(dir2):
    os.remove(os.path.join(dir2, f2))
    
myfile1 = "D:\Innovation Team\Text Search Utility\Output\Output.xlsx"
if os.path.isfile(myfile1):
    os.remove(myfile1)
    
myfile2 = "D:\Innovation Team\Text Search Utility\Output\Summary.txt"
if os.path.isfile(myfile2):
    os.remove(myfile2)
    
##################################################
# Below This is the Part, Wherein we are extracting the list of all admin users from password protected excel sheet.

temp = io.BytesIO()

with open(r"D:\Innovation Team\Text Search Utility\Setup\LIST.xlsx", 'rb') as f:
    excel = msoffcrypto.OfficeFile(f)
    excel.load_key('Shalin$Vishal')
    excel.decrypt(temp)

Data_Frame = pd.read_excel(temp)
del temp
Admin_Users = list(Data_Frame["Asset Number"])

###################################################
# Here, we are finding Assetcode from CMD and command in CMD is "hostname"

Asset_code = sp.getoutput('hostname')

###################################################
# Here, searching for asset name in the list of admin users, If asset code found in admin user list then permit to run the application or else end the program.

if Asset_code in Admin_Users:
    
    row = Data_Frame[Data_Frame['Asset Number']==Asset_code]
    Asset_Number_Name = list(row['Name'])[0]
    
    print()
    print("==========================================================")
    print()
    print("HELLO '"+Asset_Number_Name+"', YOU ARE AN AUTHORISED USER ")
    print()
    print("==========================================================")
    print()
    
    # Normally Run the program.
    
    Search_file = open(r"D:\Innovation Team\Text Search Utility\Search.txt", 'r', encoding="utf8")
    
    print()
    print("NOTE 1 :- This Application supported files types and their extensions are :- ")
    print("-------------------------------------------------------------------------------------------------------")
    print()
    print("1. Images [.png, .jpg, .jpeg, .jfif] ")
    print("2. PDF    [.pdf = Searchable as well as Scanned]") 
    print("3. Text   [.text]")
    print("4. Word   [.docx] -> Here, .doc is not supported, So, please save the file from .doc (Microsoft 97-2003) -> .docx (word document) and then again run the application")
    print()
    print()
    print("NOTE 2 :- The Expected Accurcacy Percentage(%) of different files are mentioned below :- ")
    print("-------------------------------------------------------------------------------------------------------")
    print()
    print("|-----------------------------------------------|")
    print("|     FILE TYPES       |          ACCURACY      |")
    print("|-----------------------------------------------|")
    print("|        Text          |          100 %         |")
    print("|     PDF, Word        |   Approximately  90 %  |") 
    print("|       Images         |   Approximately  80 %  |") 
    print("|-----------------------------------------------|")
    print()

    A = []
    for line in Search_file:
        A.append(line.replace("\n","").lower().strip())

    print()
    print("NOTE 3 :- Below is the List of Keywords program will search :- ")
    print("-------------------------------------------------------------------------------------------------------")
    print()  
    print(A)
    print()
    print("THANK YOU FOR YOUR TIME !!")
    print("PROGRAM HAS STARTED TO EXECUTE, CURRENTLY IN PROCESSING MODE.....")
    print()

    path = r"D:\Innovation Team\Text Search Utility\Input"  # path of Scanned Images.

    os.chdir(path) 

    Count_List = []     # List created to capture the Count of word findings from all the scanned files.

    def read_text_file(file_path):       # here, we are reading  all the scanned images one by one and extracting text from each.

        pytesseract.pytesseract.tesseract_cmd = r"D:\Innovation Team\Text Search Utility\Setup\Tesseract-OCR\tesseract.exe"

        # Load image, grayscale, and Otsu's threshold
        image = cv2.imread(file_path)
        gray = cv2.cvtColor(image, cv2.COLOR_BGR2GRAY)
        thresh = cv2.threshold(gray, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)[1]

        # Remove horizontal lines
        horizontal_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (50,1))
        detect_horizontal = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, horizontal_kernel, iterations=2)
        cnts = cv2.findContours(detect_horizontal, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        for c in cnts:
            cv2.drawContours(thresh, [c], -1, (0,0,0), 2)

        # Remove vertical lines
        vertical_kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (1,15))
        detect_vertical = cv2.morphologyEx(thresh, cv2.MORPH_OPEN, vertical_kernel, iterations=2)
        cnts = cv2.findContours(detect_vertical, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        for c in cnts:
            cv2.drawContours(thresh, [c], -1, (0,0,0), 3)

        # Dilate to connect text and remove dots
        kernel = cv2.getStructuringElement(cv2.MORPH_RECT, (10,1))
        dilate = cv2.dilate(thresh, kernel, iterations=2)
        cnts = cv2.findContours(dilate, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        cnts = cnts[0] if len(cnts) == 2 else cnts[1]
        for c in cnts:
            area = cv2.contourArea(c)
            if area < 500:
                cv2.drawContours(dilate, [c], -1, (0,0,0), -1)

        # Bitwise-and to reconstruct image
        result = cv2.bitwise_and(image, image, mask=dilate)
        result[dilate==0] = (255,255,255)

        # OCR
        data = pytesseract.image_to_string(result, lang='eng',config='--psm 6')
        #data = data.lower()
        f= open("D:\\Innovation Team\\Text Search Utility\\Output\\Converted Text Files\\"+File_Name+"="+Extensions+".txt","w+", encoding="utf8")
        f.write(data)
        f.close()
        #print("############################################")

    D = 0
    List_of_Extensions = []
    for file in os.listdir():  
        
        Split = os.path.splitext(file)
        List_of_Extensions.append(Split[1])
        
        if file.endswith(".png") or file.endswith(".jfif") or file.endswith(".jpg") or file.endswith(".jpeg") or file.endswith(".docx") or file.endswith(".pdf") or file.endswith(".PDF") or file.endswith(".txt"):   # only images with these extensions will be picked for opeartion

            D += 1
            file_path = f"{path}\{file}"
            I = file.index(".")
            File_Name = file[:I]
            Extensions = file[I+1:].upper() + " File"

            if file.endswith(".pdf"):

                with fitz.open(file_path) as doc:
                    Content = ""
                    for page in doc:
                        Content += page.get_text()

                if Content == "":
                    
                    images = convert_from_path(file_path, 500, poppler_path=r'D:\Innovation Team\Text Search Utility\Setup\poppler-0.68.0\bin')

                    New_folder_Path = r"D:\Innovation Team\Text Search Utility\Setup\Images"         
                    if not os.path.exists(New_folder_Path):
                        os.mkdir(New_folder_Path)

                    for i, image in enumerate(images):
                        image.save("D:\\Innovation Team\\Text Search Utility\\Setup\\Images\\"+str(i)+".png", "PNG")

                    pytesseract.pytesseract.tesseract_cmd = r"D:\Innovation Team\Text Search Utility\Setup\Tesseract-OCR\tesseract.exe"
                    TESSDATA_PREFIX = r"D:\Innovation Team\Text Search Utility\Setup\Tesseract-OCR"
                    tessdata_dir_config = '--tessdata-dir "D:\\Innovation Team\\Text Search Utility\\Setup\\Tesseract-OCR\\tessdata"'

                    merger = PdfFileMerger()

                    Outline = r"D:\\Innovation Team\\Text Search Utility\\Output\\Converted Searchable PDF\\"+File_Name+"= Converted to Searchable PDF"+".pdf"
                    
                    f = open(Outline, "ab")

                    for file in os.listdir(New_folder_Path):  

                        if file.endswith(".png"):        
                            filepath = os.path.join(New_folder_Path, file)

                            with Image.open(filepath) as img:

                                result = pytesseract.image_to_pdf_or_hocr(img, lang="eng", config=tessdata_dir_config)
                                pdf_file_in_memory = io.BytesIO(result)        
                                merger.append(pdf_file_in_memory)

                    merger.write(Outline)
                    merger.close()
                    f.close()
                    shutil.rmtree(New_folder_Path)    # Here, we are removing the directotries which are created after conversion it.

                    with fitz.open(Outline) as doc:
                        Content = ""
                        for page in doc:
                            Content += page.get_text()

                    f = open("D:\\Innovation Team\\Text Search Utility\\Output\\Converted Text Files\\"+File_Name+"="+Extensions+".txt","w+", encoding="utf-8")  # Made new text file 
                    f.write(Content)         # write the text extract into newly created text file.
                    f.close()
                    
                else:
                    f = open("D:\\Innovation Team\\Text Search Utility\\Output\\Converted Text Files\\"+File_Name+"="+Extensions+".txt","w+", encoding="utf-8")  # Made new text file 
                    f.write(Content)         # write the text extract into newly created text file.
                    f.close()

            elif file.endswith(".txt"):

                with open(file_path,'r') as firstfile, open(r"D:\\Innovation Team\\Text Search Utility\\Output\\Converted Text Files\\"+File_Name+"="+Extensions+".txt",'w', encoding="utf8") as secondfile:
                    for line in firstfile:
                        line = line.lower()
                        secondfile.write(line)
                
            elif file.endswith(".docx"):
                
                my_text = docx2txt.process(file_path)

                f= open("D:\\Innovation Team\\Text Search Utility\\Output\\Converted Text Files\\"+File_Name+"="+Extensions+".txt", "w+", encoding="utf8")
                f.write(my_text)
                f.close()

            else :
                read_text_file(file_path)
    
    ###########################################################################
    # Here, we are genearting "Sheet-2 = Summary"
    
    Count = len(List_of_Extensions)
    
    res = {}
    for i in List_of_Extensions:
        res[i] = List_of_Extensions.count(i)
    
    res['Total Input Files Count'] = Count

    D = []
    E = []
    F = []
    
    for key, val in res.items():
        D.append(key)
        E.append(val)
        if key == ".doc":
            F.append("Failure (This Extensions is not supported. So, Output will not be generted of this files. To Generate the output, please save this file in .docx (word document) format and then re-run the application)")
        elif key == "Total Input Files Count":
            F.append("-")            
        else:
            F.append("Success")
            
    df2 = pd.DataFrame({'Extensions': D, 'Total Count': E, 'Execution':F})   # Dataframe for Sheet-2 Generated Successfully
    
    ###########################################################################

    """  HERE, WE ARE RUNNING HALF THE PROGRAM, AS GENERATION OF TEXT FILES IS ALREADY DONE  """
    """   PART 2 :-  CORRECT CODE FOR MULTIPLE FILES OUTPUT"""

    path_text_file = r"D:\Innovation Team\Text Search Utility\Output\Converted Text Files"  # path of all text files we created.

    os.chdir(path_text_file) 

    df_final = pd.DataFrame() 

    # Here we are creating main data frame and then furthur we will just add columns.
    # In main dataframe we limited the number of rows to be the maximun count of "word to find". 
    # If value in one column less then Max_rows number, then other empty space will be filled with "NAN" values. " This will not give different length of columns error."
    
    def read_text_file(file_path):               # here, we are reading the text files one by one.
        Output = open(file_path, 'r', encoding="utf-8")

        line_number = 0
        list_of_results = []

        for line in Output:
                # For each line, check if line contains the string
            line = line.lower()
            
            line_number += 1
            if string_to_search in line:       # If yes, then add the line number & line as a tuple in the list

                COUNT = line.count(string_to_search)
                list_of_results.append((File_name, string_to_search, line_number, COUNT, line.rstrip()))

        if list_of_results == []:
            list_of_results.append((File_name, string_to_search, 0, 0, "Element Not found"))

        #print("Personal Output => " +string_to_search+ " :- ", list_of_results)
        #print()

        # Till here, we are just printing 4 differnt values into ONE LIST == "list_of_results"
        # These, "list_of_results" will contain OUTPUT of one ONE FILE.

        return list_of_results

    D = 0
    final_list = []
    for file in os.listdir():  
        
        os.listdir(path)
        
        if file.endswith(".txt"):
            D += 1
            file_path = f"{path_text_file}\{file}"

            I = file.index(".")
            File_name = file[:I]

            Combined = []
            for t in A:

                string_to_search = t
                #print("-->>  String_to_search :- ", string_to_search)
                #print()
                one_string_to_search_output = read_text_file(file_path)    # Here, This "results" contains the OUTPUT of one ONE FILE == "list_of_results."

                Combined.extend(one_string_to_search_output)
            #print("------>>>   FINAL LIST OUTPUT FOR ONE 1 FILE ")
            #print()
            #print(Combined)
            #print()
            final_list.extend(Combined)    
            print()
            print("Processing and operation on file :- '" + File_name + "' Completed")
            print()
     
    ###############################################################
    # Here, we are generating "Sheet-1 = Final Output" 

    df1 = pd.DataFrame(final_list)    # Here, making dataframe of "final_list"
    delimiter = ","                  # Swprating all the 4 elements from one block of the list into separte columns using "delimeter"
    df1[0].str.split(delimiter, expand=True)    
    df1.columns =['File_Name', 'Element_to_Find', 'Line#', 'Count', 'Full_Line']    # Naming all the 5 columns
    
    df1 = df1[df1['Line#'] != 0]          # Here, we are deleting all rows which contains 0 in "Line#" column

    ##################################################################
    # Here, we are generating "Sheet-3 = Search Value Count"
    
    df3 = df1.groupby(["Element_to_Find"]).Count.sum().reset_index()
    df3.columns = ["Search value", "Count"]
    df3.loc[len(df3.index)] = ["TOTAL", df3["Count"].sum()]    # Dataframe for Sheet-3 Generated Successfully
    
    ##################################################################
    # Here, we are generating "Sheet-4 = Search Value to file Count"

    df4 = df1.groupby(["Element_to_Find","File_Name"]).Count.sum().reset_index()
    df4.columns = ["Search value", "File_Name", "Sum of Count"]

    ##################################################################
    # Setting all the text align into all 4 dataframes
    
    df1_Final = df1.style.set_properties(**{'text-align': 'left'})
    df2_Final = df2.style.set_properties(**{'text-align': 'left'})
    df3_Final = df3.style.set_properties(**{'text-align': 'left'})
    df4_Final = df4.style.set_properties(**{'text-align': 'left'})
    
    with pd.ExcelWriter(r"D:\Innovation Team\Text Search Utility\Output\Output.xlsx") as writer:
        df1_Final.to_excel(writer, sheet_name="Final Output", index=False)
        df2_Final.to_excel(writer, sheet_name="Summary", index=False)
        df3_Final.to_excel(writer, sheet_name="Search Value Count", index=False)
        df4_Final.to_excel(writer, sheet_name="Search Value to File Count", index=False)

    print()
    print(" EXECUTION OF PROGRAM COMPLETED ")
    print()
    print("#######  THANK YOU  ############")
    print()
    
else:
    print("YOU ARE NOT AUTHORISED USER TO RUN THIS APPLICATION")
    print()


# In[4]:


import os
os.startfile(r"D:\Innovation Team\Text Search Utility\Utility\TextSearchUtility_V1.exe")

