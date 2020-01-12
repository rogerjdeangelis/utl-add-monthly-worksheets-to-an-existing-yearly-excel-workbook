# utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook
Add monthly worksheets to an existing yearly excel workbook
    Add monthly worksheets to an existing yearly excel workbook;                                                                      
                                                                                                                                      
    Output workbook (february was added to existing yearly workbook)                                                                  
    https://tinyurl.com/v6ybhqz                                                                                                       
    https://github.com/rogerjdeangelis/utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook/blob/master/year.xlsx          
                                                                                                                                      
    github                                                                                                                            
    https://tinyurl.com/yx4cgkn2                                                                                                      
    https://github.com/rogerjdeangelis/utl-add-monthly-worksheets-to-an-existing-yearly-excel-workbook                                
                                                                                                                                      
    Enhanced from here                                                                                                                
    https://tinyurl.com/tfahbfd                                                                                                       
    https://www.geeksforgeeks.org/python-how-to-copy-data-from-one-excel-sheet-to-another/                                            
                                                                                                                                      
      THREE SOLUTIONS                                                                                                                 
                                                                                                                                      
        a. Simple SAS libname (requires named ranges(tables) - avoids very messy quoting?)                                            
        b. Python with macro wrapper                                                                                                  
        c. Python without macro wrapper                                                                                               
                                                                                                                                      
      Algorithm (python more flexible?)                                                                                               
                                                                                                                                      
        1) Import openpyxl library as xl.                                                                                             
        2) Open the new month excel workbook(worksheet=FEB) using the path in which it is located.                                    
           Note: The path should be a string and have double backslashes (\\)                                                         
           instead of single backslash (\). Eg: Path should be                                                                        
            C:\\Users\\Desktop\\source.xlsx Instead of C:\Users\Admin\Desktop\source.xlsx                                             
        3) Get the name of the worksheet. We assume just the first worksheet has the new months data.                                 
                                                                                                                                      
        4) Open the yearly workbook with worksheet, JAN, in this case.                                                                
           The index of worksheet ‘n’ is ‘n-1’. For example, the index of worksheet 1 is 0.                                           
                                                                                                                                      
        5) Open the destination yearly excel workbook.                                                                                
                                                                                                                                      
        6) Calculate the total number of rows and columns in yearly excel workbook.                                                   
                                                                                                                                      
        7) Use two for loops (one for iterating through rows and another for                                                          
           iterating through columns of the excel file) to read the cell value in source file to a variable                           
           and then write it to a cell in destination file from that variable.                                                        
        8) Save the destination file.                                                                                                 
                                                                                                                                      
                                                                                                                                      
    copysheet macro on end.  Macro also in here.                                                                                      
    https://tinyurl.com/y9nfugth                                                                                                      
    https://github.com/rogerjdeangelis/utl-macros-used-in-many-of-rogerjdeangelis-repositories                                        
                                                                                                                                      
                                                                                                                                      
    I made this repo private, while I try to fix it. Used to wor?                                                                     
    utl-copy-excel-sheets-from-one-workbook-to-another-using-powershell                                                               
    https://tinyurl.com/utmr5gd                                                                                                       
    https://github.com/rogerjdeangelis/utl-copy-excel-sheets-from-one-workbook-to-another-using-powershell                            
                                                                                                                                      
    *_                   _                                                                                                            
    (_)_ __  _ __  _   _| |_                                                                                                          
    | | '_ \| '_ \| | | | __|                                                                                                         
    | | | | | |_) | |_| | |_                                                                                                          
    |_|_| |_| .__/ \__,_|\__|                                                                                                         
            |_|                                                                                                                       
    ;                                                                                                                                 
                                                                                                                                      
    %utlfkil(d:/xls/year.xlsx);  * delete if exist;                                                                                   
    %utlfkil(d:/xls/feb.xlsx);                                                                                                        
                                                                                                                                      
    * make workbook and sheets;                                                                                                       
    libname xlsa "d:/xls/year.xlsx";                                                                                                  
    libname xlsb "d:/xls/feb.xlsx";                                                                                                   
                                                                                                                                      
    data xlsa.jan(where=(sex="M")) xlsb.feb(where=(sex="F")) ;                                                                        
      set sashelp.class;                                                                                                              
    run;quit;                                                                                                                         
                                                                                                                                      
    libname xlsa clear;                                                                                                               
    libname xlsb clear;                                                                                                               
                                                                                                                                      
                                                                                                                                      
    Have yearly workbook with Jan data need to add Feb data from second workbook                                                      
                                                                                                                                      
     d:/xls/YEAR.xlsx                                                                                                                 
                                                                                                                                      
       +--------------------------------------+                                                                                       
       |  A    |  B    |  C    |  D    |  E   |                                                                                       
       +--------------------------------------+                                                                                       
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                                       
       +-------+-------+-------+-------+------|                                                                                       
     2 |Alfred |14     |M      |69     |112.5 |                                                                                       
       +-------+-------+-------+-------+------+                                                                                       
     3 |Carol  |13     |F      |56.5   |84    |                                                                                       
       ---------------------------------------+                                                                                       
       [JAN]                                                                                                                          
                                                                                                                                      
     d:/xls/FEB.xlsx                                                                                                                  
                                                                                                                                      
       +--------------------------------------+                                                                                       
       |  A    |  B    |  C    |  D    |  E   |                                                                                       
       +--------------------------------------+                                                                                       
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                                                                       
       +-------+-------+-------+-------+------|                                                                                       
     2 |Alice  |14     |F      |62.8   |102.5 |                                                                                       
       +-------+-------+-------+-------+------+                                                                                       
     3 |Henry  |14     |M      |63.5   |88    |                                                                                       
       ----------------------------------------                                                                                       
       [FEB]                                                                                                                          
                                                                                                                                      
    *            _               _                                                                                                    
      ___  _   _| |_ _ __  _   _| |_                                                                                                  
     / _ \| | | | __| '_ \| | | | __|                                                                                                 
    | (_) | |_| | |_| |_) | |_| | |_                                                                                                  
     \___/ \__,_|\__| .__/ \__,_|\__|                                                                                                 
                    |_|                                                                                                               
    ;                                                                                                                                 
                                                                                                                                      
     d:/xls/YEAR.xlsx                                                                                                                 
                                                                                                                                      
       +--------------------------------------+    +--------------------------------------+                                           
       |  A    |  B    |  C    |  D    |  E   |    |  A    |  B    |  C    |  D    |  E   |                                           
       +--------------------------------------+    +--------------------------------------+                                           
     1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|  1 |NAME   |AGE    |SEX    |HEIGHT |WEIGHT|                                           
       +-------+-------+-------+-------+------|    +-------+-------+-------+-------+------|                                           
     2 |Alfred |14     |M      |69     |112.5 |  2 |Alice  |14     |F      |62.8   |102.5 |                                           
       +-------+-------+-------+-------+------+    +-------+-------+-------+-------+------+                                           
     3 |Carol  |13     |F      |56.5   |84    |  3 |Henry  |14     |M      |63.5   |88    |                                           
       ---------------------------------------+    ----------------------------------------                                           
       [JAN]                                       [FEB]                                                                              
                                                                                                                                      
    *          _       _   _                                                                                                          
     ___  ___ | |_   _| |_(_) ___  _ __  ___                                                                                          
    / __|/ _ \| | | | | __| |/ _ \| '_ \/ __|                                                                                         
    \__ \ (_) | | |_| | |_| | (_) | | | \__ \                                                                                         
    |___/\___/|_|\__,_|\__|_|\___/|_| |_|___/                                                                                         
                   _                 _                                                                                                
      __ _     ___(_)_ __ ___  _ __ | | ___   ___  __ _ ___                                                                           
     / _` |   / __| | '_ ` _ \| '_ \| |/ _ \ / __|/ _` / __|                                                                          
    | (_| |_  \__ \ | | | | | | |_) | |  __/ \__ \ (_| \__ \                                                                          
     \__,_(_) |___/_|_| |_| |_| .__/|_|\___| |___/\__,_|___/                                                                          
                              |_|                                                                                                     
    ;                                                                                                                                 
                                                                                                                                      
    * MAKE DATA AGAIN FOR RERUNS;                                                                                                     
                                                                                                                                      
    %utlfkil(d:/xls/year.xlsx);  * delete if exist;                                                                                   
    %utlfkil(d:/xls/feb.xlsx);                                                                                                        
                                                                                                                                      
    * make workbook and sheets;                                                                                                       
    libname xlsa "d:/xls/year.xlsx";                                                                                                  
    libname xlsb "d:/xls/feb.xlsx";                                                                                                   
                                                                                                                                      
    data xlsa.jan(where=(sex="M")) xlsb.feb(where=(sex="F")) ;                                                                        
      set sashelp.class;                                                                                                              
    run;quit;                                                                                                                         
                                                                                                                                      
    libname xlsa clear;                                                                                                               
    libname xlsb clear;                                                                                                               
                                                                                                                                      
    *SOLUTION;                                                                                                                        
                                                                                                                                      
    libname xlsa "d:/xls/year.xlsx";                                                                                                  
    libname xlsb "d:/xls/feb.xlsx";                                                                                                   
                                                                                                                                      
    proc sql;                                                                                                                         
       * get name of tab in new month "FEB" in this case;                                                                             
       select                                                                                                                         
          memname into :nam trimmed                                                                                                   
       from                                                                                                                           
          sashelp.vtable                                                                                                              
       where                                                                                                                          
          libname="XLSB" and index(memname,'$')=0                                                                                     
    ;                                                                                                                                 
       create                                                                                                                         
          table xlsa.&nam as                                                                                                          
       select                                                                                                                         
          *                                                                                                                           
       from                                                                                                                           
          xlsb.&nam                                                                                                                   
    ;quit;                                                                                                                            
                                                                                                                                      
    libname xlsa clear;                                                                                                               
    libname xlsb clear;                                                                                                               
                                                                                                                                      
    *_                                                                                                                                
    | |__     _ __  _   _  __      ____ _ _ __                                                                                        
    | '_ \   | '_ \| | | | \ \ /\ / / _` | '_ \                                                                                       
    | |_) |  | |_) | |_| |  \ V  V / (_| | |_) |                                                                                      
    |_.__(_) | .__/ \__, |   \_/\_/ \__,_| .__/                                                                                       
             |_|    |___/                |_|                                                                                          
    ;                                                                                                                                 
                                                                                                                                      
    * create data for rerunning;                                                                                                      
                                                                                                                                      
    %utlfkil(d:/xls/year.xlsx);  * delete if exist;                                                                                   
    %utlfkil(d:/xls/feb.xlsx);                                                                                                        
                                                                                                                                      
    * make workbook and sheets;                                                                                                       
    libname xlsa "d:/xls/year.xlsx";                                                                                                  
    libname xlsb "d:/xls/feb.xlsx";                                                                                                   
                                                                                                                                      
    data xlsa.jan(where=(sex="M")) xlsb.feb(where=(sex="F")) ;                                                                        
      set sashelp.class;                                                                                                              
    run;quit;                                                                                                                         
                                                                                                                                      
    libname xlsa clear;                                                                                                               
    libname xlsb clear;                                                                                                               
                                                                                                                                      
    * solution;                                                                                                                       
                                                                                                                                      
    * macro on end;                                                                                                                   
    %copysheet(d:\\xls\\feb.xlsx,d:\\xls\\year.xlsx);                                                                                 
    *                                                                                                                                 
      ___   _ __  _   _   _ __   ___   __      ___ __ __ _ _ __                                                                       
     / __| | '_ \| | | | | '_ \ / _ \  \ \ /\ / / '__/ _` | '_ \                                                                      
    | (__ _| |_) | |_| | | | | | (_) |  \ V  V /| | | (_| | |_) |                                                                     
     \___(_) .__/ \__, | |_| |_|\___/    \_/\_/ |_|  \__,_| .__/                                                                      
           |_|    |___/                                   |_|                                                                         
    ;                                                                                                                                 
                                                                                                                                      
    * create data again for rerunning;                                                                                                
                                                                                                                                      
    %utlfkil(d:/xls/year.xlsx);  * delete if exist;                                                                                   
    %utlfkil(d:/xls/feb.xlsx);                                                                                                        
                                                                                                                                      
    * make workbook and sheets;                                                                                                       
    libname xlsa "d:/xls/year.xlsx";                                                                                                  
    libname xlsb "d:/xls/feb.xlsx";                                                                                                   
                                                                                                                                      
    data xlsa.jan(where=(sex="M")) xlsb.feb(where=(sex="F")) ;                                                                        
      set sashelp.class;                                                                                                              
    run;quit;                                                                                                                         
                                                                                                                                      
    libname xlsa clear;                                                                                                               
    libname xlsb clear;                                                                                                               
                                                                                                                                      
    * solution;                                                                                                                       
                                                                                                                                      
    %let wb_year=d:\\xls\\year.xlsx;  * has previous monthly worksheets ie JAN,FEB...;                                                
    %let wb_feb=d:\\xls\\feb.xlsx;  * has new monthly sorksheet, FEB in this example;                                                 
                                                                                                                                      
    %utl_submit_py64_37("                                                                                                             
    import openpyxl;                                                                                                                  
    filename ='&wb_feb';                                                                                                              
    wb1 = openpyxl.load_workbook(filename);                                                                                           
    ws1 = wb1.worksheets[0];                                                                                                          
    month = wb1.sheetnames;                                                                                                           
    filename1 ='&wb_year';                                                                                                            
    wb2 = openpyxl.load_workbook(filename1);                                                                                          
    ws2 = wb2.create_sheet(month[0]);                                                                                                 
    mr = ws1.max_row;                                                                                                                 
    mc = ws1.max_column;                                                                                                              
    for i in range (1, mr + 1):;                                                                                                      
    .   for j in range (1, mc + 1):;                                                                                                  
    .       c = ws1.cell(row = i, column = j);                                                                                        
    .       ws2.cell(row = i, column = j).value = c.value;                                                                            
    wb2.save(str(filename1));                                                                                                         
    ");                                                                                                                               
    *                          _               _                                                                                      
      ___ ___  _ __  _   _ ___| |__   ___  ___| |_                                                                                    
     / __/ _ \| '_ \| | | / __| '_ \ / _ \/ _ \ __|                                                                                   
    | (_| (_) | |_) | |_| \__ \ | | |  __/  __/ |_                                                                                    
     \___\___/| .__/ \__, |___/_| |_|\___|\___|\__|                                                                                   
              |_|    |___/                                                                                                            
    ;                                                                                                                                 
                                                                                                                                      
    %macro copysheet(wb_month,wb_year)                                                                                                
           / des="Add monthy worksheets to an existing yearly excel workbook";                                                        
                                                                                                                                      
      %utl_submit_py64_37("                                                                                                           
      import openpyxl;                                                                                                                
      filename ='&wb_month';                                                                                                          
      wb1 = openpyxl.load_workbook(filename);                                                                                         
      ws1 = wb1.worksheets[0];                                                                                                        
      month = wb1.sheetnames;                                                                                                         
      filename1 ='&wb_year';                                                                                                          
      wb2 = openpyxl.load_workbook(filename1);                                                                                        
      ws2 = wb2.create_sheet(month[0]);                                                                                               
      mr = ws1.max_row;                                                                                                               
      mc = ws1.max_column;                                                                                                            
      for i in range (1, mr + 1):;                                                                                                    
      .   for j in range (1, mc + 1):;                                                                                                
      .       c = ws1.cell(row = i, column = j);                                                                                      
      .       ws2.cell(row = i, column = j).value = c.value;                                                                          
      wb2.save(str(filename1));                                                                                                       
      ");                                                                                                                             
                                                                                                                                      
    %mend copysheet;                                                                                                                  
                                                                                                                                      
    %copysheet(d:\\xls\\feb.xlsx,d:\\xls\\year.xlsx);                                                                                 
                                                                                                                                      
                                                                                                                                      
