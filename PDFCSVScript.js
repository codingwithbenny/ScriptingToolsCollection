const fs = require("fs")
const pdfparse = require("pdf-parse")
const ObjectsToCsv = require("objects-to-csv")
const pdffile = fs.readFileSync("Name of your PDF and/or variable.pdf")

//This Script will convert a pdf into a readable csv file

pdfparse(pdffile).then(function (data) {
  //This is the array to find that address of each practice
  const suffixArray = [
    "Dr.",
    "St.",
    "Way",
    "Hwy.",
    "Ave.",
    "Cir.",
    "Blvd.",
    "Rd.",
    "Highway",
    "Lane",
    "Ct.",
    "PO Box",
    "Ln.",
    "Ridge",
    "Pkwy.",
  ]

  //This is an array to find if string is an email
  const emailCheck = ["@", ".com", ".net", ".org"]

  //The entire pdf for is taken and split based on returns
  const splitDataRough = data.text.split("\n")

  //Clean array with data in order
  const splitDataClean = []

  //Final Array with objects of each individual
  const finalArray = []

  //This for loop goes through the data that is only split by returns
  for (let i = 0; i < splitDataRough.length; i++) {
    //Looks to see if the string contains an MD, this will contain their name
    if (splitDataRough[i].includes("MD")) {
      splitDataClean.push(splitDataRough[i])
    }

    //This looks for the line that contains phone number
    if (
      splitDataRough[i].includes("-") &&
      !splitDataRough[i].includes(".") &&
      (splitDataRough[i].includes("Fax") || splitDataRough[i].length === 8)
    ) {
      //splits the phone and fax number
      const splitNumbers = splitDataRough[i].split(" ")

      //grabs the actual phone number
      const mainPhoneNumber = splitNumbers[0]
      splitDataClean.push(mainPhoneNumber)
    }

    //Checks to see if the data contains any attributes within the emailCheck array to find emails
    if (
      emailCheck.some(function (x) {
        return splitDataRough[i].indexOf(x) >= 0
      })
    ) {
      splitDataClean.push(splitDataRough[i])
    }

    //Checks to see if the data contains any attributes within the suffix array to find addresses
    if (
      suffixArray.some(function (x) {
        return splitDataRough[i].indexOf(x) >= 0
      })
    ) {
      splitDataClean.push(splitDataRough[i])
      splitDataClean.push(splitDataRough[i + 1])
    }
  }

  //This for loop will go through the clean array of the data in order and organize into objects
  for (let i = 0; i < splitDataClean.length; i++) {
    //First Grabs a name
    if (splitDataClean[i].includes("MD")) {
      let splitDataNum = ""
      let splitDataEmail = ""

      //If practice does not have a phone number listed, "" is replaced
      if (!splitDataClean[i + 3].includes("-")) {
        splitDataNum = ""
      } else {
        splitDataNum = splitDataClean[i + 3]
      }

      //If practice does not have a email listed, "" is replaced
      if (!splitDataClean[i + 4].includes("@")) {
        splitDataEmail = ""
      } else {
        splitDataEmail = splitDataClean[i + 4]
      }

      //creation of the final object
      const objectGroup = {
        Name: splitDataClean[i],
        Street: splitDataClean[i + 1],
        City: splitDataClean[i + 2],
        Number: splitDataNum,
        Email: splitDataEmail,
      }
      //Pushes each object into the final array
      finalArray.push(objectGroup)
    }
  }

  //Converts the Object array to csv
  const csv = new ObjectsToCsv(finalArray)

  //Saves CSV to pc
  csv.toDisk("./DesiredName.csv")
})
