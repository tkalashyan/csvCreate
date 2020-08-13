var rules = [
    ["ու", "u"],
    ["ա", "a"],
    ["բ", "b"],
    ["գ", "g"],
    ["դ", "d"],
    [" ե", " ye"],
    ["ե", "e"],
    ["զ", "z"],
    ["է", "e"],
    ["ը", "y"],
    ["թ", "t"],
    ["ժ", "j"],
    ["ի", "i"],
    ["լ", "l"],
    ["խ", "x"],
    ["ծ", "ts"],
    ["կ", "k"],
    ["հ", "h"],
    ["ձ", "dz"],
    ["ղ", "x"],
    ["ճ", "tch"],
    ["մ", "m"],
    ["յ", "y"],
    ["ն", "n"],
    ["շ", "sh"],
    [" ո", " vo"],
    ["ո", "o"],
    ["չ", "ch"],
    ["պ", "p"],
    ["ջ", "dj"],
    ["ռ", "vo"],
    ["ս", "s"],
    ["վ", "v"],
    ["տ", "t"],
    ["ր", "r"],
    ["ց", "c"],
    ["փ", "p"],
    ["ք", "q"],
    ["և", "ev"],
    ["օ", "o"],
    ["ֆ", "f"],
    ["․", "."]
];
var rate = 7.46;
document.getElementById("rate").value = rate;
document.getElementById("rate").addEventListener("change", rateChange);

function rateChange() {
    rate = document.getElementById("rate").value;
}

function replaceChars(string) {
    for (var i = 0; i < string.length; i++) {
        for (var j = 0; j < rules.length; j++) {
            string = string.replace(new RegExp(rules[j][0], 'gi'), rules[j][1]);
        }
    }
    return string;
}

let cust = {};
let invoiceRowI = "";
let invoiceRowH = "";
let invoiceRow = "G\tST\tINVOIC\t" + GUID + "\t932030251\tY30251_001\tDAILY\t" + StatisticsPeriod + "\t" + ReportDate + "\tSENT\r\n";
let invrptRow = "G\tST\tINVRPT\t" + GUID + "\t932030251\tY30251_001\tDAILY\t" + StatisticsPeriod + "\t" + ReportDate + "\tSENT\r\n";

function createCSV() {
    // let csvContent = "data:text/csv;charset=utf-8,"
    for (var i = 0; i < customers.length; i++) {
        var chainCode = "";
        var addressType = "";
        if (staff.includes(customers[i]["code"])) {
            chainCode = "926000010";
            addressType = "09"
        } else if (customers[i]["code"] == onlineShop) {
            chainCode = "925000212";
            addressType = "YE"
        } else {
            chainCode = "926000011";
            addressType = "16"
        }
        cust[customers[i]["code"]] = "H\t" + customers[i]["code"] + "\tAM\tArmenia\t" + customers[i]["city"] + "\t-\t" + customers[i]["realAddress"] + "\t\t" + customers[i]["name"] + "\t" + customers[i]["tin"] + "\t" + chainCode + "\t" + addressType + "\r\n"
        // invoiceRow += "H\t"+customers[i]["code"]+"\tAM\tArmenia\t-\t-\t"+customers[i]["realAddress"]+"\t\t"+customers[i]["name"]+"\t"+customers[i]["tin"]+"\t"+chainCode+"\t"+addressType+"\r\n"
    }
    if (returnItems.length) {
        for (i = 0; i < returnItems.length; i++) {
            if (items[returnItems[i]["itemNumber"]]) {
                var promoActionCode = returnItems[i]["discountPercent"] ? "DISCT" : "REGR";
                console.log(returnItems[i]["itemNumber"]);
                var initialSum = returnItems[i]["quantity"] * items[returnItems[i]["itemNumber"]]["initialPrice"];
                // console.log(i, returnItems[i]]['itemNumber']);
                if (cust[returnItems[i]["customerNumber"]]) {
                    invoiceRowH += cust[returnItems[i]["customerNumber"]];
                    delete cust[returnItems[i]["customerNumber"]];
                }
                invoiceRowI += "I\t" + PartnerMainWarehouseCode + "\t" + returnItems[i]["saleNumber"] + "\t" + returnItems[i]["saleDate"] + "\t" + returnItems[i]["customerNumber"] + "\tRET\t" + returnItems[i]["itemNumber"] + "\t" + promoActionCode + "\tNR\t-" + items[returnItems[i]['itemNumber']]['shtrix'] + "\t" + returnItems[i]['quantity'] + "\tUN\t" + initialSum + "\t" + returnItems[i]["discountPrice"] + "\tAMD\r\n"
            }
        }
    }
    if (soldItems.length) {
        for (i = 0; i < soldItems.length; i++) {
            var promoActionCode = soldItems[i]["discountPercent"] ? "DISCT" : "REGR";
            var documentNumber = soldItems[i]["documentNumber"];
            if (sales[documentNumber]) {
                for (var j = 0; j < sales[documentNumber].length; j++) {
                    var item = sales[documentNumber][j];
                    if (items[item["itemCode"]]) {
                        var shtrix = items[item["itemCode"]]["shtrix"];
                    } else {
                        console.log("Absent in ProductsRems.xlsx " + item["itemCode"]);
                        continue;
                    }
                    var initialSum = Math.round(((items[item["itemCode"]]["initialPrice"] * item["itemQuantity"]) / rate) * 100) / 100;
                    if (cust[soldItems[i]["customerNumber"]]) {
                        invoiceRowH += cust[soldItems[i]["customerNumber"]];
                        delete cust[soldItems[i]["customerNumber"]];
                    }
                    var thisMaterialCode = "";
                    if (shtrix && materialCode[shtrix]) {
                        thisMaterialCode = materialCode[shtrix]
                    } else {
                        console.log("Штрих код " + shtrix + " неизвестен.")
                    }
                    invoiceRowI += "I\t" + PartnerMainWarehouseCode + "\t" + soldItems[i]["saleNumber"] + "\t" + soldItems[i]["saleDate"] + "\t" + soldItems[i]["customerNumber"] + "\tST\t" + item["itemCode"] + "\t" + promoActionCode + "\tNR\t" + thisMaterialCode + "\t" + shtrix + "\t" + item['itemQuantity'] + "\tUN\t" + initialSum + "\t" + item["totalWithVAT"] + "\tAMD\r\n";
                }
            } else {
                console.log("Absent in SalesAnalyse.xlsx " + documentNumber);
                continue;
            }
        }
    }
    if (imports.length) {
        var invoiceNumber = imports[0]["deliveryNumber"];
        for (i = 0; i < imports.length; i++) {
            var itemCode = eans[imports[i]["ean"]] ? eans[imports[i]["ean"]]["code"] : "";
            invoiceRowI += "I\t" + PartnerMainWarehouseCode + "\t" + invoiceNumber + "\t" + imports[i]["date"] + "\t0\tLOR\t" + itemCode + "\tREGR\tNR\t" + imports[i]['productNumber'] + "\t" + imports[i]["ean"] + "\t" + imports[i]["quantity"] + "\tUN\t" + imports[i]["totalWithoutVAT"] + "\t" + Math.round((imports[i]["totalWithVAT"] * rate) * 100) / 100 + "\tAMD\r\n";
        }
    }

    var xhttp = new XMLHttpRequest();
    xhttp.open("POST", "/save-csv-string", true);
    xhttp.setRequestHeader("Content-Type", "application/json;charset=UTF-8");
    var data = {invoiceRowI: invoiceRowI, invoiceRowH: invoiceRowH, invrptRow: invrptRow}
    xhttp.send(JSON.stringify(data));

    // const itemsCode = Object.keys(items);
    // const itemsValues = Object.values(items);
    // for (var i = 0; i<itemsValues.length; i++) {
    //   var sum = 0;
    //   if (itemsValues[i]["shtrix"]) {
    //   sum = itemsValues[i]["quantity"] * itemsValues[i]["price"];
    //   initialSum = itemsValues[i]["quantity"] * itemsValues[i]["initialPrice"];
    //   invrptRow += "I\t" + itemsValues[i]["day"] + "\t" + PartnerMainWarehouseCode + "\t" + itemsCode[i] + "\tNR\t\t" + itemsValues[i]["shtrix"] + "\t" + itemsValues[i]["quantity"] + "\t" + itemsValues[i]["quantity"] + "\tUN\t" + initialSum + "\t" + sum + "\tAMD" + "\r\n"
    //   } else {
    //   console.log(itemsCode[i], itemsValues[i])
    //   }
    // }

}

function startNewDay() {
    document.getElementById(thatDay).classList.add("done");
    document.getElementById("items").classList.remove("done");
    document.getElementById("initialPrice").classList.remove("done");
    document.getElementById("price").classList.remove("done");
    document.getElementById("salesAnalyse").classList.remove("done");
    document.getElementById("sales").classList.remove("done");
    document.getElementById("return").classList.remove("done");
    document.getElementById("import").classList.remove("done");
    document.getElementById("partnerCode").classList.remove("done");
    thatDay = "";
    fileName = "";
    customers = [];
    soldItems = [];
    returnItems = [];
    imports = [];
    sales = {};
    items = {};
    eans = {};
    cust = {};
}

function finalCreateCSV() {
    invoiceRow += invoiceRowH;
    invoiceRow += invoiceRowI;
    download("ST_INVOIC_" + PartnerCode + "_" + GUID + ".txt", invoiceRow);
    download("ST_INVRPT_" + PartnerCode + "_" + GUID + ".txt", invrptRow);
}

function download(filename, text) {
    var element = document.createElement('a');
    element.setAttribute('href', 'data:text/plain;charset=utf-8,' + encodeURIComponent(text));
    element.setAttribute('download', filename);
    element.style.display = 'none';
    document.body.appendChild(element);
    element.click();
    document.body.removeChild(element);
}

function ExcelDateToJSDate(serial) {
    var utc_days = Math.floor(serial - 25569);
    var utc_value = utc_days * 86400;
    var date_info = new Date(utc_value * 1000);
    var formatedDate = date_info.getFullYear().toString();
    formatedDate += (date_info.getMonth() + 1).toString().length === 1 ? "0" + (date_info.getMonth() + 1).toString() : (date_info.getMonth() + 1).toString();
    formatedDate += date_info.getDate().toString().length === 1 ? "0" + date_info.getDate().toString() : date_info.getDate().toString();
    return formatedDate;
}

function dayFromTitle(date) {
    const d = "20" + date[2].substring(0, 2) + date[1] + date[0].substring(date[0].length - 2, date[0].length);
    console.log(d);
    return d;
}

function invrpt() {
    var shtrixs = '';
    const itemsCode = Object.keys(items);
    const itemsValues = Object.values(items);
    console.log(itemsValues);
    for (var i = 0; i < itemsValues.length; i++) {
        var sum = 0;
        if (itemsValues[i]["shtrix"] && materialCode[itemsValues[i]["shtrix"]]) {
            sum = itemsValues[i]["quantity"] * itemsValues[i]["price"];
            // console.log(itemsValues[i]["quantity"], itemsValues[i]["price"]);
            // var thisMaterialCode = "";
            // if (materialCode[itemsValues[i]["shtrix"]]) {
            const thisMaterialCode = materialCode[itemsValues[i]["shtrix"]];
            // } else {
            //   console.log("Штрих код " + itemsValues[i]["shtrix"] + " неизвестен.")
            //   shtrixs += itemsValues[i]["shtrix"] + ' ,'
            // }
            const initialSum = Math.floor(itemsValues[i]["quantity"] * itemsValues[i]["initialPrice"] * 100) / 100;
            invrptRow += "I\t" + itemsValues[i]["day"] + "\t" + PartnerMainWarehouseCode + "\t" + itemsCode[i] + "\tNR\t" + thisMaterialCode + "\t" + itemsValues[i]["shtrix"] + "\t" + itemsValues[i]["quantity"] + "\t" + itemsValues[i]["quantity"] + "\tUN\t" + initialSum + "\t" + sum + "\tAMD" + "\r\n"
        } else {
            // console.log(itemsCode[i], itemsValues[i])
        }
    }
    console.log(shtrixs);
    // items = {}
}

var fileName = "";
var customers = [];
var soldItems = [];
var returnItems = [];
var imports = [];
var sales = {};
var items = {};
var eans = {};
var materialCode = {};
var cities = ["Kotayk", "Aragacotn", "Ararat", "Arcax", "Gexarkunik", "Shirak", "Syunik", "Yerevan", "Armavir", "Armavir", "Kotayk", "Lori"]

var X = XLSX;
var XW = {
    /* worker message */
    msg: 'xlsx',
    /* worker scripts */
    worker: './xlsxworker.js'
};

var global_wb;
let thatDay = "";

var process_wb = (function () {
    // var OUT = document.getElementById('out');

    var to_json = function to_json(workbook) {
        var result = {};
        workbook.SheetNames.forEach(function (sheetName) {
            var roa = X.utils.sheet_to_json(workbook.Sheets[sheetName], {header: 1});
            if (roa.length) result[sheetName] = roa;
        });
        return JSON.stringify(result, 2, 2);
    };

    return function process_wb(wb) {
        global_wb = wb;
        var output = to_json(wb);

        // if(OUT.innerText === undefined) OUT.textContent = output;
        // else OUT.innerText = output;
        output = JSON.parse(output);
        console.log(fileName);
        switch (fileName) {
            case "Trio_SI.xlsx" :
                // $("#days").append("<p id='" + thatDay + "'>thatDay</p>");
                for (var i = 1; i < output['Trio_SI'].length - 1; i++) {
                    materialCode[output['Trio_SI'][i][5]] = output['Trio_SI'][i][4];
                }
                document.getElementById("materialCode").classList.add("done");
                console.log("materialCode", materialCode);
                break;
            case "ProductsRems.xlsx" :
                thatDay = dayFromTitle(output['Sheet1'][0][0].split("/"));
                var node = document.createElement("p");                 // Create a <li> node
                node.setAttribute("id", thatDay);
                var textnode = document.createTextNode(thatDay);         // Create a text node
                node.appendChild(textnode);                              // Append the text to <li>
                document.getElementById("days").appendChild(node);
                // $("#days").append("<p id='" + thatDay + "'>thatDay</p>");
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    items[output['Sheet1'][i][0]] = {
                        name: replaceChars(output['Sheet1'][i][2]),
                        shtrix: output['Sheet1'][i][1],
                        quantity: output['Sheet1'][i][4],
                        day: thatDay
                    };
                    eans[output['Sheet1'][i][1]] = {
                        name: replaceChars(output['Sheet1'][i][2]),
                        code: output['Sheet1'][i][0],
                        quantity: output['Sheet1'][i][4],
                        day: thatDay
                    }
                }
                document.getElementById("items").classList.add("done");
                console.log("items", items);
                console.log("eans", eans);
                break;
            case "PricesChange (2).xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    var price = Math.round((output['Sheet1'][i][4].toString().replace(/ /g, '').replace(',', '.') / rate) * 100) / 100;
                    if (items[output['Sheet1'][i][1].toString()])
                        items[output['Sheet1'][i][1].toString()]["initialPrice"] = price;
                    else
                        items[output['Sheet1'][i][0]] = {
                            name: replaceChars(output['Sheet1'][i][2]),
                            shtrix: "",
                            quantity: 0,
                            initialPrice: price,
                            day: thatDay
                        }
                }
                document.getElementById("initialPrice").classList.add("done");
                console.log("items", items);
                break;
            case "PricesChange.xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    var price = Math.round((output['Sheet1'][i][4].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100;
                    if (items[output['Sheet1'][i][1].toString()])
                        items[output['Sheet1'][i][1].toString()]["price"] = price;
                    else
                        items[output['Sheet1'][i][0]] = {
                            name: replaceChars(output['Sheet1'][i][2]),
                            shtrix: "",
                            quantity: 0,
                            initialPrice: 0,
                            price: price,
                            day: thatDay
                        }
                }
                document.getElementById("price").classList.add("done");
                console.log("items", items);
                invrpt();
                break;
            case "SalesAnalyse.xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    var saleDate = ExcelDateToJSDate(output['Sheet1'][i][0]);
                    var actionCode = output['Sheet1'][i][8] ? "DISCT" : "REGR";
                    var documentNumber = output['Sheet1'][i][1];
                    if (sales[documentNumber]) {
                        sales[documentNumber].push({
                            saleDate: saleDate,
                            itemCode: output['Sheet1'][i][3],
                            itemQuantity: output['Sheet1'][i][6],
                            itemPrice: output['Sheet1'][i][9],
                            totalWithVAT: output['Sheet1'][i][10],
                            totalWithoutVAT: "",
                            actionCode: actionCode
                        })
                    } else {
                        sales[documentNumber] = [{
                            saleDate: saleDate,
                            itemCode: output['Sheet1'][i][3],
                            itemQuantity: output['Sheet1'][i][6],
                            itemPrice: output['Sheet1'][i][9],
                            totalWithVAT: output['Sheet1'][i][10],
                            totalWithoutVAT: "",
                            actionCode: actionCode
                        }]
                    }
                }
                document.getElementById("salesAnalyse").classList.add("done");
                console.log("sales", sales);
                break;
            case "Sales.xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    var saleDate = ExcelDateToJSDate(output['Sheet1'][i][1]);
                    soldItems.push({
                        documentNumber: output['Sheet1'][i][0],
                        saleDate: saleDate,
                        saleNumber: output['Sheet1'][i][3] ? replaceChars(output['Sheet1'][i][3]) : output['Sheet1'][i][0],
                        actionType: "sale",
                        customerNumber: output['Sheet1'][i][4],
                        customerName: replaceChars(output['Sheet1'][i][5]),
                        priceType: output['Sheet1'][i][8],
                        priceWithoutDiscount: Math.round((output['Sheet1'][i][9].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        discountPercent: output['Sheet1'][i][10],
                        discountPrice: Math.round((output['Sheet1'][i][11].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        finalPrice: Math.round((output['Sheet1'][i][12].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        VAT: Math.round((output['Sheet1'][i][13].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100
                    })
                }
                document.getElementById("sales").classList.add("done");
                console.log("soldItems", soldItems);
                break;
            case "ProductsOpsByCust.xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    var returnDate = ExcelDateToJSDate(output['Sheet1'][i][0]);
                    // var saleDate = ExcelDateToJSDate(output['Sheet1'][i][2]);
                    returnItems.push({
                        saleDate: returnDate,
                        saleNumber: output['Sheet1'][i][1],
                        actionType: "return",
                        itemNumber: output['Sheet1'][i][5],
                        itemName: replaceChars(output['Sheet1'][i][6]),
                        quantity: output['Sheet1'][i][8],
                        price: Math.round((output['Sheet1'][i][9].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        discountPercent: output['Sheet1'][i][10],
                        discountPrice: Math.round((output['Sheet1'][i][12].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        customerNumber: output['Sheet1'][i][2],
                        customerName: replaceChars(output['Sheet1'][i][3])
                    })
                }
                document.getElementById("return").classList.add("done");
                console.log("returnItems", returnItems);
                break;
            case "OOO__Trio_Prodjekt__TRIO_PRODJ.xls" :
                var data = output["Sheet1"];
                console.log(data[0]);
                for (var i = 1; i < data.length - 2; i++) {
                    imports.push({
                        id: data[i][6],
                        deliveryNumber: data[i][5],
                        date: ExcelDateToJSDate(data[i][4]),
                        productNumber: data[i][7],
                        // productName : productName,
                        quantity: data[i][14] / 1000,
                        price: Math.round((data[i][15].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        totalWithVAT: Math.round((data[i][19].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        totalWithoutVAT: Math.round((data[i][16].toString().replace(/ /g, '').replace(',', '.')) * 100) / 100,
                        ean: data[i][25],
                        actionType: "import"
                    })
                }
                document.getElementById("import").classList.add("done");
                console.log("imports", imports);
                break;
            case "Customers.xlsx" :
                for (var i = 2; i < output['Sheet1'].length - 1; i++) {
                    customers.push({
                        code: output['Sheet1'][i][0],
                        name: replaceChars(output['Sheet1'][i][1]),
                        regAddress: replaceChars(output['Sheet1'][i][8]),
                        realAddress: replaceChars(output['Sheet1'][i][9]),
                        city: cities[parseInt(output['Sheet1'][i][3])] ? cities[parseInt(output['Sheet1'][i][3])] : "Yerevan",
                        tin: output['Sheet1'][i][2] ? output['Sheet1'][i][2] : output['Sheet1'][i][20]
                    })
                }
                document.getElementById("partnerCode").classList.add("done");
                console.log("customets", customers);
                break;
        }
    };
})();

var do_file = (function () {
    var rABS = typeof FileReader !== "undefined" && (FileReader.prototype || {}).readAsBinaryString;
    var domrabs = document.getElementsByName("userabs")[0];
    if (!rABS) domrabs.disabled = !(domrabs.checked = false);

    var use_worker = typeof Worker !== 'undefined';
    var domwork = document.getElementsByName("useworker")[0];
    if (!use_worker) domwork.disabled = !(domwork.checked = false);

    var xw = function xw(data, cb) {
        var worker = new Worker(XW.worker);
        worker.onmessage = function (e) {
            switch (e.data.t) {
                case 'ready':
                    break;
                case 'e':
                    console.error(e.data.d);
                    break;
                case XW.msg:
                    cb(JSON.parse(e.data.d));
                    break;
            }
        };
        worker.postMessage({d: data, b: rABS ? 'binary' : 'array'});
    };

    return function do_file(files) {
        rABS = domrabs.checked;
        use_worker = domwork.checked;
        var f = files[0];
        var reader = new FileReader();
        reader.onload = function (e) {
            var data = e.target.result;
            if (!rABS) data = new Uint8Array(data);
            if (use_worker) xw(data, process_wb);
            else process_wb(X.read(data, {type: rABS ? 'binary' : 'array'}));
        };
        if (rABS) reader.readAsBinaryString(f);
        else reader.readAsArrayBuffer(f);
    };
})();

(function () {
    var xlf = document.getElementById('xlf');
    if (!xlf.addEventListener) return;

    function handleFile(e) {
        do_file(e.target.files);
        fileName = e.target.files[0]['name'];
    }

    xlf.addEventListener('change', handleFile, false);
})();
