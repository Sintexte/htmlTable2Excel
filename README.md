# htmlTable2Excel

Do you have a Html file with alot of tables with multiple rows and collumns ?

Do you wanna convert them to excel sheets ? this code is for you .


<br>

# Installation:

Download Nodejs/Npm: https://nodejs.org/ (Tested: Nodejs v16.5)


## Install Dependencies:
#### Used Thirdparty Dependencies [jsdom, excel4node, yargs]
#### Used Dependencies [fs]

```bash
npm i
```
<br>

# Commands

## Help command:
```bash
node Htmltable2Excelsheet.js -h
```


## Command Format:
```bash
node Htmltable2Excelsheet.js -f [htmlfilelocation:REQUIRED] -e [excelfilename] 
```

## Examples:
```bash
node Htmltable2Excelsheet.js -f ./file.html -e Genetics 
```

```bash
node Htmltable2Excelsheet.js -f ./file2.html
```
