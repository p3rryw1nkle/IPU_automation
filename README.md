# IPU_automation
### To create IPUs you only need to run the 'writeSpreadsheet.py' script, however be sure to read the following points:

### 1. Make sure you download the license information spreadsheet and place it in the 'spreadsheets' folder. Rename it to 'Licenses.xlsx'. The program will look for this file to get the IPU information.

### 2. You must have a 'spreadsheets' folder in the base directory and within that folder a 'completed' folder. In the 'completed' folder there should also be two folders, one called 'email' and one 'without_email'. 

### 2. There should also be a 'logs' folder in the base directory to store the conflict.log, which will contain information anomalies for each IPU after you have run the script.

### 3. At the bottom of writeSpreadsheet.py where it says "makeFiles.process_files(initials="LB")" you will need to put in your personal initials to process your individually assigned IPUs

### 4. Any delete_me.txt files inside folders can be deleted

### 5. fixSpreadsheet may be used if there's any IPUs that need to be fixed. Place the IPUs that need to be fixed in spreadsheets/need_fixed folder, then edit fixSpreadsheet so it edits them as needed. The fixed spreadsheets will be written to spreadsheets/fixed, which you can then upload and replace the old ones.
