# TDC EEA Capacity Report

This program generates the capacity report for EEA of TDC. 


## Installation

The program is written in python and it requires it to be able to run. Follow the below sections in order to install the interpretor and required libraries. 

### Interpretor

1. Go to [https://www.python.org/](https://www.python.org/). 
2. Click "Downloads", then "Windows".
3. In the Stable Releases category, click the first version that appears. 
4. Scroll down to the Files section, and click on "Windows installer (64-bit)".
5. Right-click on the executable, and click "Run as administrator". 
6. Click on "Customize installation". 
7. In the "Optional Features"
    1. Check "pip". 
    2. Click next.
8. In the "Advanced Options"
    1. Check "Associate files with python"
    2. Check "Create shortcuts for installed applications"
    3. Check "Add Python to environment variables"
    4. Check "Install for all users"


### Crawler

On your system you'll need to have installed Microsoft Edge. The program uses it to crawl Zabbix. 


### Code

1. In the code directory, double-click on `load_venv.cmd`. 
2. Open Command prompt as administrator. 
3. Execute `python3 -m venv <path>` where `<path>` is the location of the code directory.  
4. Install libraries:
    1. In the code directory, double-click on `load_venv.cmd`. 
    2. `pip install pdfkit`
    3. `pip install pytesseract`


### PDFKit Library
1. Go to [https://wkhtmltopdf.org/index.html](https://wkhtmltopdf.org/index.html)
2. Click "Downloads".
3. Find Windows section, and click on 64-bit. 
4. Right-click on the executable, and click "Run as administrator". 
5. Install with default options.


### Tesseract Library
1. Go to [https://tesseract-ocr.github.io/](https://tesseract-ocr.github.io/)
2. Click on "User Manual".
3. Click on "Binaries".
4. Click on "Windows - Tesseract at UB Mannheim".
5. Click on the latest 64-bit installer. 
6. Right-click on the executable, and click "Run as administrator". 
7. Install with default options.


## Execution
There are two ways to execute the program. 

### Automatic
In the code directory, double-click on `execute.cmd`. 

### Manual
Open a command prompt in the code directory. 
Execute: `python main.py`.


Once the program is executed, you'll see a new instance of Microsoft Edge running. It will close itself automatically once the process is completed. 


## Description

This program allows the capacity report to generated automatically. It takes data from Zabbix graphs generated from the screen configuration Capacity Report (Razvan - Part 1), Capacity Report (Razvan - Part 2), and Capacity Report (Razvan - Part 3). 

1. Log in to Zabbix ([https://10.219.145.172:444/zabbix/zabbix.php?action=dashboard.view](https://10.219.145.172:444/zabbix/zabbix.php?action=dashboard.view))
2. Go to the Screens button
3. Goes to `Capacity Report (Razvan - Part 1)`
4. Sets the time period. 
5. Takes each image and saves it to the directory 
6. Creates a .pdf containing all graphs. 
7. Go through each downloaded image and get all the text from it. 
8. Apply corrections to the data. 
9. Get title and value from the corrected data. 
10. Add result to the main dictionary. 
11. Repeats from step 3. 
12. Go through the dictionary and create .xlsx file with the result.
13. Emulates outlook, and sends local mail directly to how should receive the report. 


### Output

The report is stored in the "output" directory. 

```
capacity_report_2022-03-01-2022-03-31.html
capacity_report_2022-03-01-2022-03-31.pdf
capacity_report_2022-03-01-2022-03-31.xlsx
```

### Graphics

All graphics are stored in the "images" directory. 

Each time period has the form of `YYYY-MM-DD-YYYY-MM-DD`. Which represents a start point and an end point. 

There are 119 graphics in total. their names are always numeric, and starts from `0` to `119`, with the extension `.png`. 

