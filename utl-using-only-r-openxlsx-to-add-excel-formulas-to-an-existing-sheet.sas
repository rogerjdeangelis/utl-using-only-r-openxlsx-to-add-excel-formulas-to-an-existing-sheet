%let pgm=utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet;

using only r openxlsx to add excel formulas to an existing sheet

   Problem
     Sum column 2 and column3 and put the sum in column3

github
https://tinyurl.com/2h9kvxa3
https://github.com/rogerjdeangelis/utl-using-only-r-openxlsx-to-add-excel-formulas-to-an-existing-sheet

Related Repos

  https://github.com/rogerjdeangelis/utl-adding-formulas-to-excel-worksheets
  https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
  https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
  https://github.com/rogerjdeangelis/utl_calculate_column_based_on_formula_stored_in_other_column
  https://github.com/rogerjdeangelis/utl_excel_add_formula_inplace
  https://github.com/rogerjdeangelis/utl_excel_add_formulas

/**************************************************************************************************************************/
/*                                          |                                  |                                          */
/*      INPUT (CREATE WORKSHEET)            |      PROCESS                     |               OUTPUT                     */
/*                                          |                                  |                                          */
/*                                          |                                  |                                          */
/* SHEET HAVE WORKBOOK D:/XLS/HAVE.XLSX     | %utl_rbegin;                     | SHEET HAVE WORKBOOK D:/XLS/HAVE.XLSX     */
/*                                          | parmcards4;                      |                                          */
/*   +-------------------------------+      | library(openxlsx)                |  +----------------------+-----------+    */
/*   |     A   |    B     |     C    |      | wb <-                            |  |    A  |  B   |   C   |     D     |    */
/*   +-------------------------------+      |  loadWorkbook("d:/xls/have.xlsx")|  +----------------------+-----------+    */
/* 1 |  MAKE   |   CITY   |  HIWAY   |      | df <- read.xlsx(                 | 1|  MAKE | CITY | HIWAY |           |    */
/*   +---------+----------+----------+      |   "d:/xls/have.xlsx"             |  +-------+------+-------+-----------+    */
/* 2 | Acura   |   17     |    23    |      |   ,sheet="have")                 | 2| Acura | 17   |   23  |=SUM(B2:C2)|    */
/*   +---------+----------+----------+      | num_rows<-nrow(df) + 1           |  +-------+------+-------+-----------+    */
/* 3 | Audi    |   22     |    31    |      | formula<-paste0(                 | 3| Audi  | 22   |   31  |=SUM(B3:C3)|    */
/*   +---------+----------+----------+      |     "=SUM(B"                     |  +-------+------+-------+-----------+    */
/* 4 | BMW     |   16     |    23    |      |    ,2:num_rows                   | 4| BMW   | 16   |   23  |=SUM(B4:C4)|    */
/*   +---------+----------+----------+      |    ,":C"                         |  +-------+------+-------+-----------+    */
/* 5 | Buick   |   15     |    21    |      |    ,2:num_rows                   | 5| Buick | 15   |   21  |=SUM(B5:C5)|    */
/*   +---------+----------+----------+      |    ,")")                         |  +-------+------+-------+-----------+    */
/*                                          | writeFormula(                    |                                          */
/* [HAVE]                                   |     wb                           | This is what you see when you open       */
/*                                          |    ,"have"                       |                                          */
/* libname sd1 "d:/sd1";                    |    ,formula                      |  +----------------------+-------+        */
/* options validvarname=upcase;             |    ,startCol=4                   |  |    A  |  B   |   C   |   D   |        */
/*                                          |    ,startRow=2                   |  +----------------------+-------+        */
/* data sd1.have;                           |    )                             | 1|  MAKE | CITY | HIWAY |       |        */
/*  set sashelp.cars(obs=48                 | saveWorkbook(                    |  +-------+------+-------+-------+        */
/*   keep=make mpg_city  mpg_highway);      |     wb                           | 2| Acura | 17   |   23  |   40  |        */
/*  by make;                                |    ,"d:/xls/have.xlsx"           |  +-------+------+-------+-------+        */
/*  if first.make;                          |    ,overwrite=TRUE)              | 3| Audi  | 22   |   31  |   53  |        */
/*  city=mpg_city;                          | ;;;;                             |  +-------+------+-------+-------+        */
/*  hiway= mpg_highway;                     | %utl_rend;                       | 4| BMW   | 16   |   23  |   39  |        */
/*  keep make city hiway;                   |                                  |  +-------+------+-------+-------+        */
/* run;quit;                                |                                  | 5| Buick | 15   |   21  |   36  |        */
/*                                          |                                  |  +-------+------+-------+-------+        */
/* %utl_rbegin;                             |                                  |                                          */
/* parmcards4;                              |                                  | [HAVE]                                   */
/*  library(openxlsx)                       |                                  |                                          */
/*  library(haven)                          |                                  |                                          */
/*  have<-read_sas("d:/sd1/have.sas7bdat")  |                                  |                                          */
/*  wb <- createWorkbook()                  |                                  |                                          */
/*  addWorksheet(wb, "have")                |                                  |                                          */
/*  writeData(wb, "have", have)             |                                  |                                          */
/*  saveWorkbook(wb, "d:/xls/have.xlsx"     |                                  |                                          */
/* ,overwrite = TRUE)                       |                                  |                                          */
/* ;;;;                                     |                                  |                                          */
/* %utl_rend;                               |                                  |                                          */
/*                                          |                                  |                                          */
/**************************************************************************************************************************/

/*                   _
(_)_ __  _ __  _   _| |_
| | `_ \| `_ \| | | | __|
| | | | | |_) | |_| | |_
|_|_| |_| .__/ \__,_|\__|
        |_|
*/

%utlfkil(d:/xls/have.xlsx);

libname sd1 "d:/sd1";
options validvarname=upcase;

data sd1.have;
 set sashelp.cars(obs=48
  keep=make mpg_city  mpg_highway);
 by make;
 if first.make;
 city=mpg_city;
 hiway= mpg_highway;
 keep make city hiway;
run;quit;

%utl_rbegin;
parmcards4;
 library(openxlsx)
 library(haven)
 have<-read_sas("d:/sd1/have.sas7bdat")
 wb <- createWorkbook()
 addWorksheet(wb, "have")
 writeData(wb, "have", have)
 saveWorkbook(wb, "d:/xls/have.xlsx"
,overwrite = TRUE)
;;;;
%utl_rend;

/**************************************************************************************************************************/
/*                                                                                                                        */
/*      INPUT (CREATE WORKSHEET)                                                                                          */
/*                                                                                                                        */
/*                                                                                                                        */
/* SHEET HAVE WORKBOOK D:/XLS/HAVE.XLSX                                                                                   */
/*                                                                                                                        */
/*   +-------------------------------+                                                                                    */
/*   |     A   |    B     |     C    |                                                                                    */
/*   +-------------------------------+                                                                                    */
/* 1 |  MAKE   |   CITY   |  HIWAY   |                                                                                    */
/*   +---------+----------+----------+                                                                                    */
/* 2 | Acura   |   17     |    23    |                                                                                    */
/*   +---------+----------+----------+                                                                                    */
/* 3 | Audi    |   22     |    31    |                                                                                    */
/*   +---------+----------+----------+                                                                                    */
/* 4 | BMW     |   16     |    23    |                                                                                    */
/*   +---------+----------+----------+                                                                                    */
/* 5 | Buick   |   15     |    21    |                                                                                    */
/*   +---------+----------+----------+                                                                                    */
/*                                                                                                                        */
/* [HAVE]                                                                                                                 */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
 _ __  _ __ ___   ___ ___  ___ ___
| `_ \| `__/ _ \ / __/ _ \/ __/ __|
| |_) | | | (_) | (_|  __/\__ \__ \
| .__/|_|  \___/ \___\___||___/___/
|_|
*/

%utl_rbegin;
parmcards4;
library(openxlsx)
wb <-
 loadWorkbook("d:/xls/have.xlsx")
df <- read.xlsx(
  "d:/xls/have.xlsx"
  ,sheet="have")
num_rows<-nrow(df) + 1
formula<-paste0(
    "=SUM(B"
   ,2:num_rows
   ,":C"
   ,2:num_rows
   ,")")
writeFormula(
    wb
   ,"have"
   ,formula
   ,startCol=4
   ,startRow=2
   )
saveWorkbook(
    wb
   ,"d:/xls/have.xlsx"
   ,overwrite=TRUE)
;;;;
%utl_rend;

/**************************************************************************************************************************/
/*                                                                                                                        */
/* This is what you see when you open                                                                                     */
/*                                                                                                                        */
/*  +----------------------+-------+                                                                                      */
/*  |    A  |  B   |   C   |   D   |                                                                                      */
/*  +----------------------+-------+                                                                                      */
/* 1|  MAKE | CITY | HIWAY |       |                                                                                      */
/*  +-------+------+-------+-------+                                                                                      */
/* 2| Acura | 17   |   23  |   40  |                                                                                      */
/*  +-------+------+-------+-------+                                                                                      */
/* 3| Audi  | 22   |   31  |   53  |                                                                                      */
/*  +-------+------+-------+-------+                                                                                      */
/* 4| BMW   | 16   |   23  |   39  |                                                                                      */
/*  +-------+------+-------+-------+                                                                                      */
/* 5| Buick | 15   |   21  |   36  |                                                                                      */
/*  +-------+------+-------+-------+                                                                                      */
/*                                                                                                                        */
/* [HAVE]                                                                                                                 */
/*                                                                                                                        */
/**************************************************************************************************************************/

/*
 _ __ ___ _ __   ___  ___
| `__/ _ \ `_ \ / _ \/ __|
| | |  __/ |_) | (_) \__ \
|_|  \___| .__/ \___/|___/
         |_|
*/

https://github.com/rogerjdeangelis/utl-adding-formulas-to-excel-worksheets
https://github.com/rogerjdeangelis/utl-sending-a-formula-to-excel-to-reference-a-cell-in-another-sheet
https://github.com/rogerjdeangelis/utl-using-sql-instead-the-excel-formula-language-for-solving-excel-problems-pyodbc
https://github.com/rogerjdeangelis/utl_calculate_column_based_on_formula_stored_in_other_column
https://github.com/rogerjdeangelis/utl_excel_add_formula_inplace
https://github.com/rogerjdeangelis/utl_excel_add_formulas

/*              _
  ___ _ __   __| |
 / _ \ `_ \ / _` |
|  __/ | | | (_| |
 \___|_| |_|\__,_|

*/
