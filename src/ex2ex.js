// @flow
const XLSX = require('xlsx');
const _ = require('lodash');

(function(){
    const inputSheet = ("./oldExcel/" + process.argv[2] + ".xlsx")
    
    const outputSheet = ("./newExcel/" + process.argv[3] + ".xlsx")
    
    const rawKeyArray = [
        'Klant',
        'Jaar',
        'Laag',
        'Campagne',
        'Land',
        'Doelgroep',
        'Soort Ad 1',
        'Soort Ad 2',
        'Soort Ad 3',
        'Soort Ad 4'
    ]

    const rawDataArray = inputExcelToJSON(inputSheet)    

    const parsedDataArray = editRawData(rawDataArray, rawKeyArray)

    const groupByModel = groupByModelCreater(parsedDataArray)

    Object.keys(groupByModel).map( (oKey,) => {
        groupByModel[oKey] = calculateNewObject(collapseArrayIntoObject(groupByModel[oKey]))
    })

    const addedObjectsArr = objectAdder(groupByModel) 

    let newWb = createWorkbook()
    
    newWb = addSheetToWorkbook(newWb, parsedDataArray, "Raw Data")
    
    newWb = addSheetToWorkbook(newWb, addedObjectsArr, "Data Analysis")

    exportNewWb(newWb, outputSheet)
})()

//Functions in order of appearance in above IIFE

/**
 * Take Excel workbook as input and return first sheet as an array of Json objects
 */
function inputExcelToJSON(inputSheet) {
    const workbook = XLSX.readFile(inputSheet)
    const rawDataArray = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]])
    return rawDataArray
}
/**
 * Take array of JSON objects as input and expand it with keys from the keyarray and values parsed from original excel
 */
function editRawData(rawDataArray, keyArray) {
    const newArray = []
        rawDataArray.forEach((e, index) => {
            const testObject = {}
            let campaignNameArr = e['Campaign Name'].split('_')

            if (campaignNameArr.length > 5) {
                campaignNameArr.pop()
            }

            let adSetNameArr = e['Ad Set Name'].split('_')[1]
            let adNameArr = e['Ad Name'].split('_')

            if (adNameArr.length > 1) {
                if (adNameArr[0].length <= 2 && adNameArr[1].length <= 2) {
                    adNameArr.shift()
                    adNameArr.shift()
                }
            }

            let valueArray = [...campaignNameArr, adSetNameArr, ...adNameArr]

            for (i = 0; i < keyArray.length; i++) {
                testObject[keyArray[i]] = valueArray[i]
            }

            delete e['Campaign Name']
            delete e['Ad Name']
            delete e['Ad Set Name']
            let returnObject = Object.assign({}, testObject, e)
            newArray.push(returnObject)
    })
    return newArray
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
    const groupByModel = _.groupBy(dataArray, (row) => {
      return [row['Laag'], row['Land']];
    });
  
    return groupByModel;
  }
  /**
 * Collapse array of objects into single object (There must be a better way to do this...)
 */
function collapseArrayIntoObject(array) {
    let returnObject = {
        Land: array[0]['Land'],
        Impressions: 0,
        AmountSpent: 0,
        WebsiteClicks: 0,
        WebsiteContentViews: 0,
        Purchases7: 0,
        Purchases28: 0,
        UniquePurchases: 0
    }

    array.forEach((e) => {
        returnObject.Impressions += e.Impressions
        returnObject.AmountSpent += typeParser(e['Amount Spent (EUR)'])
        returnObject.WebsiteClicks += typeParser(e['Link Clicks'])
        returnObject.WebsiteContentViews += typeParser(e['Website Content Views'])
        returnObject.Purchases7 += typeParser(e['Purchases [7 Days After Viewing]'])
        returnObject.Purchases28 += typeParser(e['Purchases [28 Days After Clicking]'])
        returnObject.UniquePurchases += (typeParser(e['Unique Purchases [7 Days After Viewing]']) + typeParser(e['Unique Purchases [28 Days After Clicking]']))
    })
   return returnObject
 }
/**
 * typeParser (checks for empty strings and converts string to numbers)
 */
function typeParser(input) {
    if (input == '') {
        return 0
    }
    else {
        return parseFloat(input)
    }
 }
 /**
 * Turn object into Data Report-friendly object for excel sheet
 */
function calculateNewObject(obj) {
    return {
        'Rijlabels': obj.Land,
        'Budget spent': obj.AmountSpent,
        'Impressions': obj.Impressions,
        'Website Clicks': obj.WebsiteClicks,
        'Website Visits': obj.WebsiteContentViews,
        'Purchases [28 Days PC]': obj.Purchases28,
        'Purchases [7 Days PI]': obj.Purchases7,
        'CPM €': ((obj.AmountSpent * 1000) / obj.Impressions),
        'CPC €': (obj.AmountSpent / obj.WebsiteClicks),
        'CTR %': (obj.WebsiteClicks / obj.Impressions),
        'PC Con %':(obj.Purchases28 / obj.WebsiteClicks),
        'PI Con %': (obj.Purchases7 / obj.Impressions),
        'Cost per landing view': (obj.AmountSpent / obj.WebsiteContentViews),
        'Cost per PC Con': (obj.AmountSpent / obj.Purchases28),
        'Cost per PI Con': (obj.AmountSpent / obj.Purchases7),
        'Som van Purch Con value (TOTAL)': (obj.Purchases28 + obj.Purchases7)
    }
}
/**
 * add Objects to arrays based on country
 */
function objectAdder(groupByModel) {
    const rtArray = [], prArray = [], awArray = []
    let result = []
    Object.keys(groupByModel).forEach( (oKey,) => {
        if (oKey.includes('RT')) {
            rtArray.push(groupByModel[oKey])
        }
        else if (oKey.includes('PR')) {
            prArray.push(groupByModel[oKey])
        }
        else if (oKey.includes('AW')) {
            awArray.push(groupByModel[oKey])
        }
    })
    const rtObject = (rtArray == []) ? {} : arrayReducer(rtArray, 'RT')
    const prObject = (prArray == []) ? {} : arrayReducer(prArray, 'PR')
    const awObject = (awArray == []) ? {} : arrayReducer(awArray, 'AW')

    return [ {}, rtObject, ...rtArray, {}, prObject, ...prArray, {}, awObject, ...awArray ]
}
/**
 * Take array of Data Report objects and add their properties. Return single object with added properties
 */
function arrayReducer(array, type) {
    return {
        'Rijlabels': type,
        'Budget spent': _.sumBy(array, (o) => { return o['Budget spent']; }),
        'Impressions': _.sumBy(array, (o) => { return o['Impressions']; }),
        'Website Clicks': _.sumBy(array, (o) => { return o['Website Clicks']; }),
        'Website Visits': _.sumBy(array, (o) => { return o['Website Visits']; }),
        'Purchases [28 Days PC]': _.sumBy(array, (o) => { return o['Purchases [28 Days PC]']; }),
        'Purchases [7 Days PI]': _.sumBy(array, (o) => { return o['Purchases [7 Days PI]']; }),
        'CPM €': _.sumBy(array, (o) => { return o['CPM €']; }),
        'CPC €': _.sumBy(array, (o) => { return o['CPC €']; }),
        'CTR %': _.sumBy(array, (o) => { return o['CTR %']; }),
        'PC Con %': _.sumBy(array, (o) => { return o['PC Con %']; }),
        'PI Con %': _.sumBy(array, (o) => { return o['PI Con %']; }),
        'Cost per landing view': _.sumBy(array, (o) => { return o['Cost per landing view']; }),
        'Cost per PC Con': _.sumBy(array, (o) => { return o['Cost per PC Con']; }),
        'Cost per PI Con': _.sumBy(array, (o) => { return o['Cost per PI Con']; }),
        'Som van Purch Con value (TOTAL)': _.sumBy(array, (o) => { return o['Som van Purch Con value (TOTAL)']; }),
    }
}
/**
 * Create new Excel workbook 
 */
 function createWorkbook () {
     const newWb = XLSX.utils.book_new()
     return newWb
}
/**
 * Take array of JSON objects as input, convert it to a sheet and return Excel workbook with sheet in it
 */
function addSheetToWorkbook (newWb, newDataArray, newSheetName) {
    const newSheet = XLSX.utils.json_to_sheet(newDataArray)
    XLSX.utils.book_append_sheet(newWb, newSheet, newSheetName)
    return newWb
}
/**
 * Export new Excel workbook to location passed in cli
 */
function exportNewWb(newWb, outputSheet) {
    XLSX.writeFile(newWb, outputSheet)
}
