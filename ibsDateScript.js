const xlsx = require("xlsx")
const ObjectsToCsv = require("objects-to-csv")
const wb = xlsx.readFile("Larkspur Employee Wages 2020 to 2021.xlsx", {
  cellDates: true,
})
const sheets = wb.SheetNames

const addDate = (sheet) => {
  //Chooses the second sheet since the first one contains nothing
  const ws = wb.Sheets[sheets[1]]

  //converts the excel data to a workable JSON
  const data = xlsx.utils.sheet_to_json(ws)

  const payrollTypes = []
  //Array with all of the raw data
  const messyData = []
  //Indexes of the split keyboard to vreak up into smaller arrays
  const splitIndexes = []
  //Array with each person and their respective data
  const cleanData = []
  //Final array of objects that can be converted to a CSV
  const finalArray = []

  let currentEmp = ""
  //Loops through every row in the Excel file
  for (let i = 0; i < data.length; i++) {
    //This grabs every value in each row and places in an array
    const allValues = Object.values(data[i])

    //console.log(allValues)

    if (allValues.length > 2) {
      //console.log(allValues[0])

      //This checks to make sure we are accessing an employee by checking their ID
      if (allValues[1].includes("DD") && allValues[1].length === 6) {
        //Adds a split to know where employees start and end
        messyData.push("splitHere")
        //Adds first instance of employee to messy data
        messyData.push(allValues)
        //Sets the current employee equal to the first instance in order to grab everytime their name pops up
        currentEmp = allValues[3]

        if (!payrollTypes.includes(allValues[4])) {
          payrollTypes.push(allValues[4])
        }
      } else if (allValues[0] === currentEmp) {
        //Pushes a row if the employee matches from the start of the row
        if (!payrollTypes.includes(allValues[1])) {
          payrollTypes.push(allValues[1])
        }
        messyData.push(allValues)
      }
    }
  }

  //Adds an ending keyword for the split function to work properly
  messyData.push("splitHere")

  //Loops through every value in messy data and finds the indexes of the keyword and adds to splitIndexes
  for (let x = 0; x < messyData.length; x++) {
    if (messyData[x] === "splitHere") {
      splitIndexes.push(x)
    }
  }

  //counter for the indexes
  let iCounter = 0

  //This takes the each index of an array and splits it at the indexes and pushes the new array to cleanData
  for (const j of splitIndexes) {
    //splits messyData at the two indexes building an array with the values we need
    cleanData.push(messyData.slice(iCounter, j))
    //adds 1 to the index of j to progress the splitting process
    iCounter = j + 1
  }

  //Loops through the cleanData array and breaks it off into values and formats
  for (let i = 0; i < cleanData.length; i++) {
    //This if statement is a catchall to make sure no empty arrays get through
    if (cleanData[i].length > 1) {
      for (let x = 0; x < cleanData[i].length; x++) {
        //grabs the date from the first array of arrays and formats to mm-dd-yyyy
        let date = new Date(cleanData[i][0][0])
        let niceDate = date.toISOString().substring(0, 10)

        //Checks if the current array is index 0 since the values for the first array are in different indexes
        if (x === 0) {
          //formats the final object
          const finalObj = {
            Date: niceDate,
            Name: cleanData[i][0][3],
          }

          for (let z = 0; z < payrollTypes.length; z++) {
            finalObj[payrollTypes[z]] = 0
          }
          const newSum = finalObj[cleanData[i][0][4]] + cleanData[i][0][6]

          finalObj[cleanData[i][0][4]] = Math.round(newSum * 100) / 100

          //pushes the final object
          finalArray.push(finalObj)
        } else {
          //formats the final object
          const finalObj = {
            Date: niceDate,
            Name: cleanData[i][x][0],
            PayrollItem: cleanData[i][x][1],
            Amount: cleanData[i][x][3],
          }

          for (let z = 0; z < payrollTypes.length; z++) {
            finalObj[payrollTypes[z]] = 0
          }

          const newSum = finalObj[cleanData[i][x][1]] + cleanData[i][x][3]

          finalObj[cleanData[i][x][1]] = Math.round(newSum * 100) / 100

          //pushes the final object
          finalArray.push(finalObj)
        }
      }
    }
  }

  //console.log(finalArray[11])

  //This is used to merge objects with the same name and date and add the sums
  let helper = {}
  let resultArray = finalArray.reduce(function (r, o) {
    //key
    let key = o.Name + "-" + o.Date

    if (!helper[key]) {
      // create a copy of the object
      helper[key] = Object.assign({}, o)
      //console.log(helper[key])
      //places the copy in the accumulator
      r.push(helper[key])
    } else {
      //for every helper, find the codeType key, grab the value that correlates, and add on to helper
      for (let i = 0; i < payrollTypes.length; i++) {
        //console.log(helper[key][payrollTypes[i]])

        const newSum = helper[key][payrollTypes[i]] + o[payrollTypes[i]]
        helper[key][payrollTypes[i]] = Math.round(newSum * 100) / 100
      }
    }

    return r
  }, [])

  //this package turns out data into a csv
  const csv = new ObjectsToCsv(resultArray)

  //the csv is then saved to your disk
  csv.toDisk("./IBSWithDates.csv")
}

addDate()
