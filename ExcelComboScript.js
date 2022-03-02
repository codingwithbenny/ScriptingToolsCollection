const xlsx = require("xlsx")

//reads the excel file and converts date to ISO format
const wb = xlsx.readFile("Name of your excel sheet and/or variable", {
  cellDates: true,
})

const sheets = wb.SheetNames

//takes all specific data from excel WB and turns into JSON
const sheetNames = function () {
  //this is the final JSON
  const finalWBObj = []
  //Grabs each individual sheet/tab in the excel WB
  for (let i = 0; i < sheets.length; i++) {
    //grabs all of the data within one sheet
    const ws = wb.Sheets[sheets[i]]
    //converts the ws data into a workable JSON
    const data = xlsx.utils.sheet_to_json(ws)

    //find the employees Name and Number within the single sheet
    const empNum = Object.keys(data[1])[0]
    const empName = Object.keys(data[1])[1]
    //Loops through each individual row of data
    for (let x = 0; x < data.length; x++) {
      //checks to see if row contains a date
      if (Object.values(data[x])[0] instanceof Date) {
        //Object that is pushed into finalWBObj to form JSON
        const finalObj = {}
        //Array that combines the employee name, number and all of the values in the row of data
        const finalArray = new Array(empNum, empName, ...Object.values(data[x]))
        //loops through the array of data and converts the array into object format

        for (let y = 0; y < finalArray.length; y++) {
          finalObj[y] = finalArray[y]
        }
        //Takes the built object and places it into the finalWBObj Array
        finalWBObj.push(finalObj)

        console.log(finalWBObj)
      }
    }
  }
  //creates a new excel wb
  const newWB = xlsx.utils.book_new()
  //converts finalWBObjJSON into sheets format
  const newWS = xlsx.utils.json_to_sheet(finalWBObj)
  //places the new sheet inside of the new workbook with a tab named "Combined Data"
  xlsx.utils.book_append_sheet(newWB, newWS, "Desired Sheet Name")

  //Takes the completed workbook and creates the excel file on PC with chosen name
  xlsx.writeFile(newWB, "DesiredName.xlsx")
  // console.log(finalWBObj)
}
//Calls the function
sheetNames(sheets)
