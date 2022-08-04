# IPU_automation
### To create IPUs you only need to run the 'writeSpreadsheet.py' script, however be sure to read the following points:

### 1. Must have a 'spreadsheets' folder in the base directory and within that folder a 'completed' folder. In the 'completed' folder there should also be two folders, one called 'email' and one 'without_email'. 

### 2. There should also be a 'logs' folder in the base directory to store the conflict.log, which contains information anomalies for each IPU

### 3. At the bottom of writeSpreadsheet.py where it says "makeFiles.process_files(initials="LB")" you will need to put in your personal initials to process your individually assigned IPUs

### 4. Also make sure to download and drag in the spreadsheet containing all license information into the 'spreadsheets' folder and rename it to "Licenses.xlsx". The program will look for this file to get the IPU information.

### 5. Any delete_me.txt files inside folders can be deleted

### 6. fixSpreadsheet may be used if there's any IPUs that need to be fixed. Place the IPUs that need to be fixed in spreadsheets/need_fixed folder, then edit fixSpreadsheet so it fixes them like how you want. The fixed spreadsheets will be written to spreadsheets/fixed, which you can then upload and replace the old ones.
