const xlsx = require("xlsx")
const ObjectsToCsv = require("objects-to-csv")
const wb = xlsx.readFile("IBSExcel.xlsx", { cellDates: true })

const sheets = wb.SheetNames

const addName = (sheet) => {
  const ws = wb.Sheets[sheets[0]]

  //converts the excel data to a workable JSON
  const data = xlsx.utils.sheet_to_json(ws)

  //An array of all different types of codes
  const codeTypes = []
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

    //If a cell in a row includes the word 'regular' or 'Quarter Reconciliation Delta' each of the values is pushed into messy data
    if (
      allValues.includes("Regular") ||
      allValues.includes("Quarter Reconciliation Delta")
    ) {
      //Date
      messyData.push(allValues[1])
      //Code
      messyData.push(allValues[4])
      //Record Amount
      messyData.push(allValues[7])

      //This checks for the code in every row to add to our array for seperation
      if (!codeTypes.includes(allValues[4])) {
        codeTypes.push(allValues[4])
      }
    }

    //console.log(codeTypes)
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

  //This is an organized array of our combined objects
  const organizeArray = []

  //This loops through each array in cleanData
  for (let y = 0; y < cleanData.length; y++) {
    //This Grabs the first and last name in each array and combines them
    let fullName = cleanData[y][0] + " " + cleanData[y][1]

    //This function loops through each value of cleanData[y]
    for (let z = 0; z < cleanData[y].length; z++) {
      //The Final object will require a name and date to reference
      let finalObj = {
        Name: fullName,
        Date: "",
      }

      //Adds each code to the object as a key and sets it equal to 0
      for (let i = 0; i < codeTypes.length; i++) {
        finalObj[codeTypes[i]] = 0
      }

      //Looks for every instance of a Date
      if (cleanData[y][z] instanceof Date) {
        //Sets the date to a usable format
        let date = new Date(cleanData[y][z])
        let newDate = date.toISOString().substring(0, 10)

        //Sets date in final Object
        finalObj["Date"] = newDate

        //looks to see if the value of the code is a number or a -
        if (cleanData[y][z + 2] === "-") {
          //if - set code to 0
          finalObj[cleanData[y][z + 1]] += 0
        } else {
          //else add the amount to the code
          finalObj[cleanData[y][z + 1]] += cleanData[y][z + 2]
        }
        //sends the object up to the organized array
        organizeArray.push(finalObj)
      }
    }
  }

  //This is used to merge objects with the same name and date and add the sums
  let helper = {}
  let result = organizeArray.reduce(function (r, o) {
    //key
    let key = o.Name + "-" + o.Date

    if (!helper[key]) {
      // create a copy of the object
      helper[key] = Object.assign({}, o)
      console.log(helper[key])
      //places the copy in the accumulator
      r.push(helper[key])
    } else {
      //for every helper, find the codeType key, grab the value that correlates, and add on to helper
      for (let i = 0; i < codeTypes.length; i++) {
        helper[key][codeTypes[i]] += o[codeTypes[i]]
      }
    }

    return r
  }, [])

  //console.log(result)

  // //this package turns out data into a csv
  const csv = new ObjectsToCsv(result)

  //the csv is then saved to your disk
  csv.toDisk("./PreferedName.csv")
}

addName()
