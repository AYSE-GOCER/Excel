### Excel Projects
Example 1:  

**Data Cleaning**, usage of **XLOOKUP** function, **INDEX** function, **MATCH** function, **IF** function, creation of **Pivot Table** and **Dashboard**

![excel_2](https://github.com/user-attachments/assets/05ebb6d1-c1b2-454c-86a1-bfa21f62b3bc)

- =XLOOKUP(C2;customers_updated!$A$1:$A$1001;customers_updated!$B$1:$B$1001;;0)
- =IF(XLOOKUP(C4;customers_updated!$A$1:$A$1001;customers_updated!$C$1:$C$1001;;0)=0;"";XLOOKUP(C4;customers_updated!$A$1:$A$1001;customers_updated!$C$1:$C$1001;;0))
- =INDEX(products_updated!$A$1:$G$49;KAÇINCI(orders_updated!$D2;products_updated!$A$1:$A$49;0);MATCH(orders_updated!I$1;products_updated!$A$1:$G$1;0))
- =IF(I9="Rob";"Robusta"; IF(I9="Exc";"Excelsa"; IF(I9="Ara";"Arabica"; IF(I9="Lib";"Liberica"))))

See the [Sample Excel File for Analysis of Coffee Sales Data](https://github.com/AYSE-GOCER/Excel/blob/main/Excel%20Project%202%20CoffeOrdersData.xlsx)

Example 2: 

**INDEX** function, creation of **drop-down list**, **specific print area** set up, **EDATE** funtion to quickly add or subtract months from a date, combining cells into one cell.

![excel_1](https://github.com/user-attachments/assets/4dbb469a-4ca2-4c0d-9a55-d97cc9cf5cba)

- =INDEX(C3:H3;$A$1)
- =EDATE(C3;1)

See the [Sample Excel File for Analysis Certificate.xlsx](https://github.com/AYSE-GOCER/Excel/blob/main/Sample%20Excel%20File%20for%20Analysis%20Certificate.xlsx)
