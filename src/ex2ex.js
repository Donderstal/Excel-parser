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

    const analysisKeyArray = [
        'Rijlabels',
        'Budget spent',
        'Impressions',
        'Website clicks',
        'Website visits',
        'Purchases (28 days)',
        'Purchases (7 days)',
        'CPM €',
        'CPC €',
        'CTR %',
        'PC Conversion rate %',
        'PI Conversion rate %',
        'Cost per landing page view',
        'Cost per PC Conversion €',
        'Cost per PI Conversion €',
        'Som van Purch Value (TOTAL)'
    ]

    const rawDataArray = inputExcelToJSON(inputSheet)    

    const parsedDataArray = editRawData(rawDataArray, rawKeyArray)

    const groupByModel = analyseParsedData(parsedDataArray, analysisKeyArray)
  // reduce the group by model into a single object 
  //  

    let newWb = createWorkbook()
    
    newWb = addSheetToWorkbook(newWb, parsedDataArray, "Raw Data")
    
    newWb = addSheetToWorkbook(newWb, [{'Test': 'Test'}, {' Test2': 'Test2'}], "Data Analysis")

    exportNewWb(newWb, outputSheet)
})()


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
 *  'PR': {
 *    'NL': [{data}, {data2}, ...]
 *    ....
 *    },
 *  ...
 *  }
 */
function analyseParsedData(dataArray, analysisKeyArray) {
  // group By Laag Key
  const groupByModel = _.groupBy(dataArray, (row) => {
    return row['Laag'] + ':' + row['Land'];
  });
  console.log(_.keys(groupByModel));

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
        returnObject.AmountSpent += mathHelper(e['Amount Spent (EUR)'])
        returnObject.WebsiteClicks += mathHelper(e['Link Clicks'])
        returnObject.WebsiteContentViews += mathHelper(e['Website Content Views'])
        returnObject.Purchases7 += mathHelper(e['Purchases [7 Days After Viewing]'])
        returnObject.Purchases28 += mathHelper(e['Purchases [28 Days After Clicking]'])
        returnObject.UniquePurchases += (mathHelper(e['Unique Purchases [7 Days After Viewing]']) + mathHelper(e['Unique Purchases [28 Days After Clicking]']))
    })
   return returnObject
 }

/**
 * Mathhelper (checks for empty strings and converts string to numbers)
 */
 function mathHelper(input) {
    if (input == '') {
        return 0
    }
    else {
        return parseFloat(input)
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
