const fs = require("fs")
const jsdom = require("jsdom")
const EXCEL = require("excel4node")
const { JSDOM } = jsdom;
var argv = require('yargs/yargs')(process.argv.slice(2))
    .usage('Usage: $0 -f [location] -e [location]')
    .demandOption(['f'])
    .alias('f', 'htmlFileLocation')
    .alias('e', 'excelFileName')
    .describe('f', 'Html file location')
    .describe('e', 'Excel file name')
    .argv;

var excelFilename = "file.xlsx"
if (argv.e)
{
    excelFilename = argv.e + ".xlsx"
}
var excelDir = __dirname + "/" + excelFilename
var workbook = new EXCEL.Workbook()
var style = workbook.createStyle({
    font: {
        color: '#000000',
        size: 16
    },
    numberFormat: '$#,##0.00; ($#,##0.00); -'
});

var htmlFile = fs.readFile(argv.f, "utf8", (err, htmlString) =>
{
    let domHtml = new JSDOM(htmlString)
    let tableElements = domHtml.window.document.querySelectorAll("table")
    let count = 0

    tableElements.forEach(tableElement =>
    {
        let sheetName = "Sheet" + count
        let workSheet = workbook.addWorksheet(sheetName)
        let trElements = tableElement.querySelectorAll("tr")
        let trCount = 1
        let tdCount = 0

        console.log("---- " + sheetName + " ----")
        trElements.forEach(trElement =>
        {
            let tdElements = trElement.querySelectorAll("td")
            tdElements.forEach(tdElement =>
            {
                let value = tdElement.textContent
                workSheet.cell(trCount, tdCount + 1).string(value).style(style)
                tdCount++
            })
            tdCount = 0
            trCount++
        })

        count++
    });

    console.log("Saving [" + excelFilename + "] ...")
    workbook.write(__dirname + '/' + excelFilename);
})