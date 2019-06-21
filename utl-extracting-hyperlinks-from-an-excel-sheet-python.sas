Extracting hyperlinks from an excel sheet python                                                                                
                                                                                                                                
Excel is a moving target so not sure this will work for all version of excel                                                    
                                                                                                                                
My setup                                                                                                                        
                                                                                                                                
Excel 2010 64bit                                                                                                                
Python 2.7 64bit                                                                                                                
SAS 9.4M6                                                                                                                       
                                                                                                                                
I don't think R can do this easily,                                                                                             
                                                                                                                                
github                                                                                                                          
https://tinyurl.com/yy32v2pp                                                                                                    
https://github.com/rogerjdeangelis/utl-extracting-hyperlinks-from-an-excel-sheet-python                                         
                                                                                                                                
Spreadsheet                                                                                                                     
https://tinyurl.com/y6arobml                                                                                                    
https://github.com/rogerjdeangelis/utl-extracting-hyperlinks-from-an-excel-sheet-python/blob/master/Roster.xlsx                 
                                                                                                                                
https://tinyurl.com/y2uros82                                                                                                    
https://communities.sas.com/t5/SAS-Programming/Reading-Excel-hyperlinks-and-stroring-them-in-a-new-column-SAS-9/m-p/567895      
                                                                                                                                
                                                                                                                                
*_                   _                                                                                                          
(_)_ __  _ __  _   _| |_                                                                                                        
| | '_ \| '_ \| | | | __|                                                                                                       
| | | | | |_) | |_| | |_                                                                                                        
|_|_| |_| .__/ \__,_|\__|                                                                                                       
        |_|                                                                                                                     
;                                                                                                                               
                                                                                                                                
d:/xls/roster.xlsx                                                                                                              
                                                                                                                                
      +----------------------------------------------------------------+-------------------------+------------+------------+    
      |     A      |    B       |     C      |    D       |    E       |    F       |    G       |    H       |    I       |    
      +----------------------------------------------------------------+-------------------------+------------+------------+    
   1  | Active Roster Report With IQ                                                                          |            |    
      +------------+------------+------------+------------+------------+------------+------------+------------+------------+    
   2  |            |            |            |            |            |            |            |            |            |    
      +------------+------------+------------+------------+------------+------------+------------+------------+------------+    
   3  |            |            |            |            |            |            |            |            |            |    
      +------------+------------+------------+------------+------------+------------+------------+------------+------------+    
   4  |            | Facility   |   Name     |  Booking   |  Location  |   IQ       |IQ Category | MH Roster  | Diagnosis  |    
      +------------+------------+------------+------------+------------+------------+------------+------------+------------+    
   5  |            |   TEST     | TEST(LINK) | 000000     | LOCATION   |  100       |   TEST     |  TEST      |   TEST     |    
      +------------+------------+------------+------------+------------+------------+------------+------------+------------+    
                                    ^                                                                                           
                                    |                                                                                           
                                  Cell C5                                                                                       
    [ROSTER TEST]             Has the Hyperlink                                                                                 
                                                                                                                                
*            _               _                                                                                                  
  ___  _   _| |_ _ __  _   _| |_                                                                                                
 / _ \| | | | __| '_ \| | | | __|                                                                                               
| (_) | |_| | |_| |_) | |_| | |_                                                                                                
 \___/ \__,_|\__| .__/ \__,_|\__|                                                                                               
                |_|                                                                                                             
;                                                                                                                               
                                                                                                                                
hl_obj.display  None                                                                                                            
hl_obj.target   https://www.sapphireemr.com/Main/Pages/PatientChartPage.aspx?pid=00000000                                       
hl_obj.tooltip  None                                                                                                            
hl_obj          ref='C5', location=None, tooltip=None, display=None, id='rId1'                                                  
                                                                                                                                
*                                                                                                                               
 _ __  _ __ ___   ___ ___  ___ ___                                                                                              
| '_ \| '__/ _ \ / __/ _ \/ __/ __|                                                                                             
| |_) | | | (_) | (_|  __/\__ \__ \                                                                                             
| .__/|_|  \___/ \___\___||___/___/                                                                                             
|_|                                                                                                                             
;                                                                                                                               
                                                                                                                                
%utl_submit_py64('                                                                                                              
import openpyxl;                                                                                                                
wb = openpyxl.load_workbook("d:/xls/roster.xlsx");                                                                              
ws = wb["TEST ROSTER"];                                                                                                         
hl_obj = ws.cell(row = 5, column = 3).hyperlink;                                                                                
if hl_obj:;                                                                                                                     
.   print(hl_obj.display);                                                                                                      
.   print(hl_obj.target);                                                                                                       
.   print(hl_obj.tooltip);                                                                                                      
.   print(hl_obj);                                                                                                              
');                                                                                                                             
