const xlsx = require("xlsx")
const ObjectsToCsv = require("objects-to-csv")
const wb = xlsx.readFile("Name of your excel sheet and/or variable.xlsx", {
  cellDates: true,
})

const sheets = wb.SheetNames

const addName = (sheet) => {
  const ws = wb.Sheets[sheets[0]]

  //converts the excel data to a workable JSON
  const data = xlsx.utils.sheet_to_json(ws)

  //Array with all of the raw data
  const messyData = []
  //Indexes of the split keyboard to vreak up into smaller arrays
  const splitIndexes = []
  //Array with each person and their respective data
  const cleanData = []
  //Final array of objects that can be converted to a CSV
  const finalArray = []

  //Loops through every row in the Excel file
  for (let i = 0; i < data.length; i++) {
    //This grabs every value in each row and places in an array
    const allValues = Object.values(data[i])

    //If a cell in a row includes the "First Name" a keyword is placed to signify a new person and the first name is pushed to messy data
    if (allValues.includes("First Name")) {
      messyData.push("newPerson")
      messyData.push(allValues[1])
    }

    //If a cell in a row includes the  "Last Name" the value is pushed to messy data
    if (allValues.includes("Last Name")) {
      messyData.push(allValues[1])
    }

    //If a cell in a row includes the word regular each of the values is pushed into messy data
    if (allValues.includes("Regular")) {
      //Date
      messyData.push(allValues[1])
      //Number
      messyData.push(allValues[3])
      //Code
      messyData.push(allValues[4])
      //Record Amount
      messyData.push(allValues[7])
      //Record Amount ER
      messyData.push(allValues[8])
    }
  }

  //this is a one time push to the messy array to make sure final data is included, this is because the function grabs the values between two keyboards
  messyData.push("newPerson")

  //Loops through every value in messy data and finds the indexes of the keyword and adds to splitIndexes
  for (let x = 0; x < messyData.length; x++) {
    if (messyData[x] === "newPerson") {
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

  //This specific file contains junk data in the beginning so I get rid of it here
  cleanData.shift()

  //This loops through each array in cleanData
  for (let y = 0; y < cleanData.length; y++) {
    //This Grabs the first and last name in each array and combines them
    let fullName = cleanData[y][0] + " " + cleanData[y][1]

    //This loops through each value in each array
    for (let z = 0; z < cleanData[y].length; z++) {
      //Here we test to see if the value is an instance of a Date, if so, we know our place
      if (cleanData[y][z] instanceof Date) {
        //These two convert the date to normal mm-dd-yyyy format
        let date = new Date(cleanData[y][z])
        let niceDate = date.toISOString().substring(0, 10)

        //Using the date index, we place each value we need in an object
        const finalObj = {
          Name: fullName,
          Date: niceDate,
          Number: cleanData[y][z + 1],
          Code: cleanData[y][z + 2],
          RecordAmount: cleanData[y][z + 3],
          RecordAmountER: cleanData[y][z + 4],
        }

        //We grab our final object and place it in the final array
        finalArray.push(finalObj)
      }
    }
  }

  console.log(finalArray)

  //this package turns out data into a csv
  const csv = new ObjectsToCsv(finalArray)

  //the csv is then saved to your disk
  csv.toDisk("./DesiredName.csv")
}

addName()
