# Excel-Shredder
Shred the gnarliest of Excel files

Step 1. Verify Python 3 is available  
  Open Terminal and run:  
    python3 --version  
  If you see “Python 3.x.x”, you’re good.  

Step 2. Ensure pip is installed  
    python3 -m ensurepip --upgrade  

Step 3. Install the required packages  
    python3 -m pip install pandas openpyxl tkinterdnd2  

Step 4. Save the shredder script  
  • In a new folder (e.g. ~/Projects/ExcelShredder)  
  • Create a file named shredder.py  
  • Paste in your Python code  

Step 5. Launch the shredder by running the code

Step 6. Shred your Excel files  
  • Drag-and-drop one or more .xls/.xlsx files onto the window  
  • Or click “Browse…” and select multiple files at once  
  • The script picks sheet 2 if it exists, otherwise sheet 1  
  • It splits data into 999-row CSV chunks  
  • It creates a new folder on your Desktop for each workbook  

Step 7. Repeat or Exit  
  • Drag more files to process additional workbooks  
  • Click “Exit” to close the window  

Intended output  
  For each dropped file you’ll get:  
    • A folder on your Desktop named after the workbook  
    • Inside that folder, CSV files named “<workbook> – start–end.csv”  
  Multiple files can be processed in one go.  

