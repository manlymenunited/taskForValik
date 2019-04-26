let osmosis = require('osmosis');
let Excel = require('exceljs');

let resultArr = [];
let workbook = new Excel.Workbook();
let sheet = workbook.addWorksheet('Sheet1');
sheet.columns = [
    { header: 'id', key: 'id' },
    { header: 'type', key: 'type' },
    { header: 'name', key: 'name' },
    { header: 'lNum', key: 'lNum' },
    { header: 'licenceNum', key: 'licenceNum' },
    { header: 'licenceDate', key: 'licenceDate' },
    { header: 'customCode', key: 'customCode' },
    { header: 'address', key: 'address' },
];

url = "https://www.alta.ru/svh/";

osmosis
    .get(url)
    .paginate("//div[@class='pagination']//tr/td[last()]/a", 83)
    .find("//div[@class='boxSubstrate boxSubstrate-offset-0 p-10']")
    .set({
        type: "text()",
        name: "div[@class='h3']",
        license: "div[@class='pSvh_fieldColumn pSvh_fieldColumn-list pSvh_fieldColumn-right']",
        address: "div[@class='pSvh_fieldColumn pSvh_fieldColumn-list pSvh_fieldColumn-left lightgray']",
        customCode: "div[@class='pSvh_fieldColumn pSvh_fieldColumn-list pSvh_fieldColumn-left']",
    })
    .log(console.log)
    .data((data) => { resultArr.push(data); })
    .done(() => {
        for (let i = 0; i < resultArr.length; i++) {
            let arrLicense = resultArr[i].license.split("\n");
            for (let k = 0; k < arrLicense.length; k++) {
                let myLicense = arrLicense[k].trim();
                sheet.addRow({
                    id: i,
                    type: resultArr[i].type.replace(/^\s+/g, ""),
                    name: resultArr[i].name.replace(/^\s+/g, ""),
                    lNum: k + 1,
                    licenceNum: myLicense.substr(0, myLicense.indexOf("действует")),
                    licenceDate: myLicense.substr(myLicense.indexOf("действует") + 12),
                    customCode: resultArr[i].customCode.replace(/^\s+/g, ""),
                    address: resultArr[i].address.replace(/^\s+/g, "")
                })
            }
        }
        workbook.xlsx.writeFile("valikTask.xlsx").then(function () { });
    })
    ;