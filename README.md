# P6-Network-Graph
A series of scripts and reports to generate a network graph of P6 nodes and edges in Gephi

![alt text](https://github.com/Smone5/P6-Network-Graph/blob/master/P6_network_schedule.png "P6 Network Graph")


## Installation 
Oracle P6, Microsoft Excel and Gephi (2010 or greater) must be installed. The Excel scripts are were designed and tested in Excel 2010. The data was only tested and used in Gephi 0.9.1 only. The reports and data came from Oracle P6 version 8.3. These instructions assume basic working knowledge of Oracle P6, Microsoft Excel and Gephi. 

## Usage
1. In Oracle P6, import the report "P6 Nodes Export Report.erp"
2. In Oracle P6, import the report "P6 Relationships Export.erp"
3. Run and save the reports "P6 Nodes Export Report" and "P6 Relationships Export" in P6. Save the reports as:
 + ASCII Text File
 + Field Delimiter = Tab
 + Text Qualifier = '
4. In Excel, open the file "flatten_edges"
5. Import the data from the CSV file Relationships Export you created into cell A1, tab delimited.
6. Run the the macro "linksp6" in the Excel file "flatten_edges" you have opened
7. Copy the new data from columns "D", "E", "F" and "G" into a new Excel workbook.
8. Save the data copied into the new Excel workbook as a CSV file and maybe call it "edges" in order to be uploaded into Gephi. 
9. Open a new Gephi file
10. Select File -> Import Spreadsheet
11. Import the "edges" CSV file you created on step 8 into Gephi's edges table:
 + Seperator: Comma
 + As table: Edges table
 + Charset: UTF-8
12. Import the "nodes" CSV file you created in step 3 into Gephi;s Node table:
  + Seperator: Tab
  + As table: Nodes table
  + UTF-8
13. Run the layout to Force Atlas or other layout to start analyzing data. 
