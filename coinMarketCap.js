// node coinMarketCap.js --source=https://coinmarketcap.com/ --dest=table.json
// npm init
// npm i minimist
// npm i jsdom

let minimist = require('minimist');
let args = minimist(process.argv);
let jsdom = require('jsdom');
let fs = require('fs');
let axios = require('axios');
let tabletojson = require('tabletojson').Tabletojson;
let puppeteer = require('puppeteer');
let XlsxPopulate = require('xlsx-populate');

automation();

function createTable(i){
    tabletojson.convertUrl(
        args.source,
        function(tablesAsJson) {
            let json = JSON.stringify(tablesAsJson[0]);
            let fileName = "table" + i + ".json";
            fs.writeFileSync(fileName, json , "utf-8");
            createExcelSheet(i , fileName );
        }
    );
}

function createExcelSheet(idx , fileName){
    // Load a new blank workbook
XlsxPopulate.fromBlankAsync(fileName)
.then(workbook => {
    // Modify the workbook.
    workbook.sheet("Sheet1").cell("A1").value("Name").style("bold", true );
    workbook.sheet("Sheet1").cell("C1").value("Price").style("bold", true);
    workbook.sheet("Sheet1").cell("E1").value("24h %").style("bold", true);
    workbook.sheet("Sheet1").cell("G1").value("7d %").style("bold", true);;
    workbook.sheet("Sheet1").cell("I1").value("Market Cap").style("bold", true);
    workbook.sheet("Sheet1").cell("L1").value("Volume 24h").style("bold", true);
    workbook.sheet("Sheet1").cell("P1").value("Circulating Supply").style("bold", true);
    let readingJsonFile = fs.readFileSync(fileName , "utf-8");
    let dataJSO = JSON.parse(readingJsonFile);
    console.log(dataJSO.length);
    for(let i = 0 ; i < dataJSO.length ; i++){
        let idx = i + 2;
        workbook.sheet("Sheet1").cell("A" + idx).value(dataJSO[i]['Name']);
        workbook.sheet("Sheet1").cell("C"+ idx).value(dataJSO[i]['Price']);
        workbook.sheet("Sheet1").cell("E"+ idx).value(dataJSO[i]['24h %']);
        workbook.sheet("Sheet1").cell("G"+ idx).value(dataJSO[i]['7d %']);
        workbook.sheet("Sheet1").cell("I"+ idx).value(dataJSO[i]['Market Cap']);
        workbook.sheet("Sheet1").cell("L"+ idx).value(dataJSO[i]['Volume(24h)']);
        workbook.sheet("Sheet1").cell("P"+ idx).value(dataJSO[i]['Circulating Supply']);
    }
    // Write to file.
    return workbook.toFileAsync(`./Bitcoin${idx}.xlsx`);
});
}

async function automation() {
    axios.get(args.source).then(function(res){
        let html = res.data;
        let dom = new jsdom.JSDOM(html);
        let doc = dom.window.document;
        createTable(0);
    })
}
automation();
