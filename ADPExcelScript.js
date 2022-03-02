const xlsx = require("xlsx")

const wb = xlsx.readFile("Name of your excel sheet and/or variable", {
  cellDates: true,
})

const sheets = wb.SheetNames

//This function takes each individual sheet inside of the excel workbook and converts it into one large array of values
const parseSheetIntoArray = (individualSheet) => {
  //calls the individual sheet
  const ws = wb.Sheets[sheets[individualSheet]]

  //converts the excel data to a workable JSON
  const data = xlsx.utils.sheet_to_json(ws)

  //This is the going to be the final, messy array
  const currentData = []

  //This loops through each of the rows of data and picks out only the data we want, it also converts some of the data to be usable
  for (let x = 0; x < data.length; x++) {
    //This grabs every piece of data in every row and sticks them into an array
    const allValues = Object.values(data[x])

    //Here we sort through every single value and filter out what we want and dont want
    for (let y = 0; y < allValues.length; y++) {
      //The current value we are testing to see if we want it or not
      const currentValue = allValues[y]

      //All the values we want happen do be strings so we are only going to search them
      if (typeof currentValue === "string") {
        //This looks for each employees name in the format Employee: Doe, John
        if (currentValue.includes("Employee:")) {
          //From here it will take the employee string, cut out the "Employee: " and add the word split in front of it before adding to our big array
          //split is an indicator when each employee stops and begins
          currentData.push("split", currentValue.substring(10))
        }

        //The creator of these tax documents made our lives harder by mixing dates, wages and a bunch of other data into one long string. We only want the dates and wages

        //This searches for the row that contains all of the data, it starts with "Check Date:"
        if (currentValue.includes("Check Date:")) {
          //Split up the string into an array
          let split = currentValue.split(" ")

          //grab the date
          let date = split[2]

          //a place for the wage once it is cleaned up
          let wage = ""

          //this looks through the array to find the wages because the data can move position unlike the date
          for (let z = 0; z < split.length; z++) {
            //each individual item of the split array
            const splitItem = split[z]

            //All wages contain a $, that is what we are looking for
            if (splitItem.includes("$")) {
              //At times the wages will contain a "," so we must get rid of it to convert to a number
              if (splitItem.includes(",")) {
                //if there is a comma, this boots it out, gets rid of the $ and sets the wage
                wage = Number(splitItem.replace(",", "").substring(1))
              } else {
                //if there is no comma, just boot out the $ and set the wage
                wage = Number(splitItem.substring(1))
              }
            }
          }

          //Places the date and wage in an array inside of our final currentData array
          currentData.push([date, wage])
        }

        //No longer need SDI
        //Looks for the value CA SDI
        // if (currentValue.includes("CA SDI")) {
        //   //Finds the index of CA SDI and since the value is next to it in the array, we add 1 to the index and grab the value
        //   const newIndex = allValues.indexOf(currentValue) + 1

        //   //we can use this new index to grab the and push it up to the currentData array in this format:  "CA SDI",123
        //   currentData.push(currentValue, allValues[newIndex])
        // }

        //Looks for the value CA SUI-ER
        if (currentValue.includes("CA SUI-ER")) {
          //Finds the index of CA SUI and since the value is next to it in the array, we add 1 to the index and grab the value
          const newIndex = allValues.indexOf(currentValue) + 1

          //we can use this new index to grab the and push it up to the currentData array in this format:  "CA SUI-ER",123
          currentData.push(currentValue, allValues[newIndex])
        }
      }
    }
  }

  //Sends the current messy array back to the seperate values function to clean it up
  return currentData
}

//This file takes the mess arrays from above and sorts them out into clean arrays, performs calculations, and creates the final excel sheet
const seperateValues = () => {
  //This is the final array of objects that will be converted into the excel sheet
  const cleanArray = []

  //The main loop that sends each individual excel sheet through the cleansing process
  for (let a = 0; a < sheets.length; a++) {
    //This is the main messy array from parseSheetIntoArray()
    const mainArray = parseSheetIntoArray(a)

    //The for loop below needs an extra split on the end, to get the final value
    mainArray.push("split")

    //The indexes of each "split" in the messy array
    const splitIndexes = []

    //An array of the messy arrays but seperated at each split
    const splitArrayDirty = []

    //This takes the mainArray, and performs the split at each occurance of the word "split", then creates a new array in splitDirtyArray
    for (let i = 0; i < mainArray.length; i++) {
      if (mainArray[i] === "split") {
        splitIndexes.push(i)
      }
    }

    let i = 0

    for (const j of splitIndexes) {
      splitArrayDirty.push(mainArray.slice(i, j))
      i = j + 1
    }

    //This is where the cleansing of the dirty arrays takes place and organizes the data
    //Looping through each array in splitArrayDirty
    for (let i = 0; i < splitArrayDirty.length; i++) {
      const currentArray = splitArrayDirty[i]

      //This counter counts the number of arrays in each array to perform an averaging calculation
      let numberOfArray = 0

      //This is the function that looks for every array in the current array which holds the date and wage of the current employee
      for (let x = 0; x < currentArray.length; x++) {
        if (Array.isArray(currentArray[x])) {
          //adds 1 to the array counter
          numberOfArray++
        }
      }

      //These calcuate the SUI and SDI Average for each date by dividing their values by the numberOfArray

      let averageSUI = Math.round((currentArray[2] / numberOfArray) * 100) / 100
      //No Longer need SDI
      // let averageSDI = Math.round((currentArray[4] / numberOfArray) * 100) / 100

      //This function puts together each of the arrays and then converts them to objects
      for (let y = 0; y < currentArray.length; y++) {
        //This if statement looks for each array which holds date and wage, it then sandwiches the date and wage between the employee name "currentArray[0]" and the averaged SUI and ADI
        if (Array.isArray(currentArray[y])) {
          //the final object that will be pushed inside cleanArray
          const finalObj = {}

          //The sandwiching of the values into one array
          const finalArray = new Array(
            currentArray[0],
            currentArray[y][0],
            currentArray[y][1],
            averageSUI
            //averageSDI
          )

          //Setting each sandwiched array into an object for JSON format
          for (let y = 0; y < finalArray.length; y++) {
            finalObj[y] = finalArray[y]
          }

          //Placing the final object inside the cleanArray
          cleanArray.push(finalObj)
        }
      }
    }
  }

  console.log(cleanArray)

  //creates a new excel wb
  const newWB = xlsx.utils.book_new()
  //converts finalWBObjJSON into sheets format
  const newWS = xlsx.utils.json_to_sheet(cleanArray)
  //places the new sheet inside of the new workbook with a tab named "Combined Data"

  //converts to excel sheet
  xlsx.utils.book_append_sheet(newWB, newWS, "Desired Name")

  //Takes the completed workbook and creates the excel file on PC with chosen name
  // xlsx.writeFile(newWB, "CombinedDataADP.xlsx")

  //Takes the completed workbook and creates a csv file on PC with chosen name
  xlsx.writeFile(newWB, "Desired Name.csv")
}

//calls the entire process
seperateValues()
