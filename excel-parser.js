const XLSX = require('xlsx');
const _ = require('lodash');

(function () {
    const inputSheet = process.argv[2];

    const outputSheet = process.argv[3];

    const rawKeyArray = ['Klant', 'Jaar', 'Laag', 'Campagne', 'Land', 'Doelgroep', 'Soort Ad 1', 'Soort Ad 2', 'Soort Ad 3', 'Soort Ad 4'];

    const rawDataArray = inputExcelToJSON(inputSheet);

    const parsedDataArray = editRawData(rawDataArray, rawKeyArray);3

    const groupByModel = groupByModelCreater(parsedDataArray);

    Object.keys(groupByModel).map(oKey => {
        const countryRows = groupByModel[oKey];
        groupByModel[oKey] = makeDataReportRow(reduceCountryRows(countryRows));
    });

    let dataReportRows = makeLaagDataRows(groupByModel);

    dataReportRows = euroAndPercentAdder(dataReportRows)

    let newWb = createWorkbook();

    newWb = addSheetToWorkbook(newWb, parsedDataArray, "Raw Data");

    newWb = addSheetToWorkbook(newWb, dataReportRows, "Data Analysis");

    exportNewWb(newWb, outputSheet);
})();
//Functions in order of appearance in above IIFE

/**
 * Take Excel workbook as input and return first sheet as an array of Json objects
 */
function inputExcelToJSON(inputSheet) {
    const workbook = XLSX.readFile(inputSheet);
    const rawDataArray = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]);
    return rawDataArray;
}
/**
 * Take array of JSON objects as input and expand it with keys from the keyarray and values parsed from original excel
 */
function editRawData(rawDataArray, keyArray) {
    const newArray = [];
    rawDataArray.forEach(e => {
        const testObject = {};
        let campaignNameArr = e['Campaign Name'].split('_');

        if (campaignNameArr.length > 5) {
            campaignNameArr.pop();
        }

        let adSetNameArr = e['Ad Set Name'].split('_')[1];
        let adNameArr = e['Ad Name'].split('_');

        if (adNameArr.length > 1) {
            if (adNameArr[0].length <= 2 && adNameArr[1].length <= 2) {
                adNameArr.shift();
                adNameArr.shift();
            }
        }

        let valueArray = [...campaignNameArr, adSetNameArr, ...adNameArr];

        for (i = 0; i < keyArray.length; i++) {
            testObject[keyArray[i]] = valueArray[i];
        }

        delete e['Campaign Name'];
        delete e['Ad Name'];
        delete e['Ad Set Name'];
        let returnObject = Object.assign({}, testObject, e);
        newArray.push(returnObject);
    });
    return newArray;
}
/**
 * Create array of JSON objects based on analysisKeyArray
 * finalModel will be something like: 
 * { 
 *  [PR,NL]: [{data}, {data2}, ...]
 *  }
 */
function groupByModelCreater(dataArray) {
    // group By Laag Key and Land Key
    const groupByModel = _.groupBy(dataArray, row => {
        return [row['Laag'], row['Land']];
    });

    return groupByModel;
}
/**
 * Collapse array of rows into single object
 */
function reduceCountryRows(array) {
    let returnObject = {
        Land: array[0]['Land'],
        Impressions: 0,
        AmountSpent: 0,
        WebsiteClicks: 0,
        WebsiteContentViews: 0,
        Purchases7: 0,
        Purchases28: 0,
        sumTotVal: 0
    };

    array.forEach(e => {
        returnObject.Impressions += e.Impressions;
        returnObject.AmountSpent += typeParser(e['Amount Spent (EUR)']);
        returnObject.WebsiteClicks += typeParser(e['Link Clicks']);
        returnObject.WebsiteContentViews += typeParser(e['Website Content Views']);
        returnObject.Purchases7 += typeParser(e['Purchases [7 Days After Viewing]']);
        returnObject.Purchases28 += typeParser(e['Purchases [28 Days After Clicking]']);
        returnObject.sumTotVal += (typeParser(e['Purchases Conversion Value [7 Days After Viewing]']) + typeParser(e['Purchases Conversion Value [28 Days After Clicking]']));
    });
    return returnObject;
}
/**
 * typeParser (checks for empty strings and converts them to zeros)
 */
function typeParser(input) {
    if (input == '') {
        return 0;
    } else {
        return parseFloat(input);
    }
}
/**
* Turn object into Data Report-friendly object for excel sheet
* takes original excel row
* returns new (calculated) excel row
*/
function makeDataReportRow(obj) {
    return {
        'Rijlabels': obj.Land,
        'Budget spent': parseFloat(obj.AmountSpent.toFixed(2)),
        'Impressions': obj.Impressions,
        'Website Clicks': obj.WebsiteClicks,
        'Website Visits': obj.WebsiteContentViews,
        'Purchases [28 Days PC]': obj.Purchases28,
        'Purchases [7 Days PI]': obj.Purchases7,
        'CPM €': (obj.AmountSpent * 1000 / obj.Impressions).toFixed(2),
        'CPC €': (obj.AmountSpent / obj.WebsiteClicks).toFixed(2),
        'CTR %': (obj.WebsiteClicks / obj.Impressions * 100).toFixed(2),
        'PC Con %': (obj.Purchases28 / obj.WebsiteClicks * 100).toFixed(2),
        'PI Con %': (obj.Purchases7 / obj.Impressions).toFixed(4),
        'Cost per landing view': (obj.AmountSpent / obj.WebsiteContentViews).toFixed(2),
        'Cost per PC Con': (obj.AmountSpent / obj.Purchases28).toFixed(2),
        'Cost per PI Con': (obj.AmountSpent / obj.Purchases7).toFixed(2),
        'Som van purchases': (obj.Purchases28 + obj.Purchases7),
        'Value per Conversie': (obj.sumTotVal / (obj.Purchases28 + obj.Purchases7)).toFixed(2),
        'Som van Purch Con value (TOTAL)': parseFloat((obj.sumTotVal).toFixed(2))
    };
}
/**
 * take groupByModel
 * Returns new excelsheet array of row objects
 */
function makeLaagDataRows(groupByModel) {
    const rtArray = [],
          prArray = [],
          awArray = [];
    Object.keys(groupByModel).forEach(oKey => {
        if (oKey.includes('RT')) {
            rtArray.push(groupByModel[oKey]);
        } else if (oKey.includes('PR')) {
            prArray.push(groupByModel[oKey]);
        } else if (oKey.includes('AW')) {
            awArray.push(groupByModel[oKey]);
        }
    });

    //
    const rtObject = rtArray.length === 0 ? {} : rowsReducer(rtArray, 'RT');
    let prObject = prArray.length === 0 ? {} : rowsReducer(prArray, 'PR');
    const awObject = awArray.length === 0 ? {} : rowsReducer(awArray, 'AW');

    rtArray.sort((a,b) => (a['Rijlabels'] > b['Rijlabels']) ? 1 : ((b['Rijlabels'] > a['Rijlabels']) ? -1 : 0))
    prArray.sort((a,b) => (a['Rijlabels'] > b['Rijlabels']) ? 1 : ((b['Rijlabels'] > a['Rijlabels']) ? -1 : 0))
    awArray.sort((a,b) => (a['Rijlabels'] > b['Rijlabels']) ? 1 : ((b['Rijlabels'] > a['Rijlabels']) ? -1 : 0))

    const totArray = [rtObject, prObject, awObject]
    const totObject =  rowsReducer(totArray, 'TOTAAL');

    let finalArray = [{}, prObject, ...prArray, {}, rtObject, ...rtArray, {}, awObject, ...awArray, {}, totObject];
    
    finalArray = finalArray.map(obj => {
        if (obj.length === 0) {
            return {}
        }
        else {
            return numberDotsAndCommas(obj)
        }
    })
    return finalArray
}
/**
 * Add dots and commas to make number easily readable
 */
function numberDotsAndCommas(rowObj) {
    return rowObj = _.mapValues(rowObj, (propVal) => { 
        if (rowObj['Rijlabels'] == propVal) {
            return propVal
        }
        else if (rowObj['PI Con %'] == propVal) {
            return propVal.replace(/[,.]/g, (m) => {
                return m === ',' ? '.' : ',';   
            });
        }
        else {
            propVal = new Intl.NumberFormat('en-NL').format(propVal) 
            propVal = propVal.replace(/[,.]/g, (m) => {
                return m === ',' ? '.' : ',';   
            }); 
            if (propVal.split(',')[1] !== undefined && propVal.split(',')[1].length < 2) {
                propVal += 0           
            }
            return propVal
        }
    });
}
/**
 * Take array of Data Report objects and add their properties. Return single object with added properties
 */
function rowsReducer(array, type) {
    const budSpent = _.sumBy(array, o => {
        return o['Budget spent'];            
    });
    const imp = _.sumBy(array, o => {
        return o['Impressions'];
    });
    const clicks = _.sumBy(array, o => {
        return o['Website Clicks'];
    });
    const visits = _.sumBy(array, o => {
        return o['Website Visits'];
    });
    const pu28 = _.sumBy(array, o => {
        return o['Purchases [28 Days PC]'];
    });
    const pu7 = _.sumBy(array, o => {
        return o['Purchases [7 Days PI]'];
    });
    const sumTotVal = _.sumBy(array, o => {
        return o['Som van Purch Con value (TOTAL)'];
    });
    const sumPurch = _.sumBy(array, o => {
        return o['Som van purchases'];
    });
    return {
        'Rijlabels': type,
        'Budget spent': parseFloat(budSpent)/* .toFixed(2) */,
        'Impressions': imp,
        'Website Clicks': clicks,
        'Website Visits': visits,
        'Purchases [28 Days PC]': pu28,
        'Purchases [7 Days PI]': pu7,
        'CPM €': (budSpent * 1000 / imp).toFixed(2),
        'CPC €': (budSpent / clicks).toFixed(2),
        'CTR %': (clicks / imp * 100).toFixed(2),
        'PC Con %': (pu28 / clicks * 100).toFixed(2),
        'PI Con %': (pu7 / imp).toFixed(4),
        'Cost per landing view': (budSpent / visits).toFixed(2),
        'Cost per PC Con': (budSpent / pu28).toFixed(2),
        'Cost per PI Con': (budSpent / pu7).toFixed(2),
        'Som van purchases': sumPurch,
        'Value per Conversie': (sumTotVal / sumPurch).toFixed(2),
        'Som van Purch Con value (TOTAL)': sumTotVal
    };
}
/**
 * Adds euro and percent marks
 */
function euroAndPercentAdder(dataReportRows) {
    return dataReportRows.map(obj => {
        if (obj.hasOwnProperty('Budget spent')) {
            return {        
                'Rijlabels': obj['Rijlabels'],
                'Budget spent': '€'+ obj['Budget spent'],
                'Impressions': obj['Impressions'],
                'Website Clicks': obj['Website Clicks'],
                'Website Visits': obj['Website Visits'],
                'Purchases [28 Days PC]': obj['Purchases [28 Days PC]'],
                'Purchases [7 Days PI]': obj['Purchases [7 Days PI]'],
                'CPM €': '€'+obj['CPM €'],
                'CPC €': '€'+obj['CPC €'],
                'CTR %': obj['CTR %'] + " %",
                'PC Con %': obj['PC Con %'] + " %",
                'PI Con %': obj['PI Con %'] + " %" ,
                'Cost per landing view': '€'+obj['Cost per landing view'],
                'Cost per PC Con': '€'+obj['Cost per PC Con'],
                'Cost per PI Con': '€'+obj['Cost per PI Con'],
                'Som van purchases': obj['Som van purchases'],
                'Value per Conversie': '€'+obj['Value per Conversie'],
                'Som van Purch Con value (TOTAL)': obj['Som van Purch Con value (TOTAL)']
            }
        }
        else {
            return {}
        }
    })
}
/**
 * Create new Excel workbook 
 */
function createWorkbook() {
    const newWb = XLSX.utils.book_new();
    return newWb;
}
/**
 * Take array of JSON objects as input, convert it to a sheet and return Excel workbook with sheet in it
 */
function addSheetToWorkbook(newWb, newDataArray, newSheetName) {
    const newSheet = XLSX.utils.json_to_sheet(newDataArray);
    XLSX.utils.book_append_sheet(newWb, newSheet, newSheetName);
    return newWb;
}
/**
 * Export new Excel workbook to location passed in cli
 */
function exportNewWb(newWb, outputSheet) {
    XLSX.writeFile(newWb, outputSheet);
}