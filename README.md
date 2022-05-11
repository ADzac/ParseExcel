# ParseExcel
Contributors :  Adam Bin Azmi
                Muhammad Aniq Daniel Bin Mohd Zakri
                Muhammad Izzat Affandi Bin Adnan
                
Taking all Excel files in a selected directory, take its filename and save it in a table.
The function 'parse' is used to be able to find all wanted information from the excel file and save it into a table
The function 'append_csv' is used to put all the info in the table to a specific column. 
The function 'saveExcel' is the main function where all two of the funtion above are used. It also open,save and close the excel file .This is crucial if the excel file was downloaded from a web/browser. As this will have window to be able to 'know' that it is dealing with an .xls file.
