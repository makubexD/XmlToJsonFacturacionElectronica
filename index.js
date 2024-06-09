const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const xlsx = require('xlsx');

const xmlFolder = './XmlFiles'; // Folder containing XML files
const outputExcelFile = './output.xlsx'; // Output Excel file

// Equivalence dictionary for element names
const equivalenceDict = {
  'cbc:ID': ['cbc:ID', 'cbc:CompanyID', 'cbc:maku', 'cbc:demo'],
  'cbc:Description': ['cbc:Description'],
  'cbc:IssueDate': ['cbc:IssueDate'],
  'cbc:DocumentCurrencyCode': ['cbc:DocumentCurrencyCode']
};

// Function to convert XML to JSON
const convertXmlToJson = (xml) => {
  return new Promise((resolve, reject) => {
    xml2js.parseString(xml, { explicitArray: false }, (err, result) => {
      if (err) {
        reject(err);
      } else {
        resolve(result);
      }
    });
  });
};

// Function to find the value of a key based on equivalence dictionary
const findValue = (obj, keys) => {
  for (const key of keys) {
    if (obj[key]) {
      return typeof obj[key] === 'object' && obj[key]['_'] ? obj[key]['_'] : obj[key];
    }
  }
  return 'Not found';
};

// Function to extract specific values from the JSON data
const extractValues = (jsonData) => {
  try {
    const identifier = Object.keys(jsonData)[0];
    let jsonResponse = {};

    switch (identifier) {
      case 'Invoice':
        jsonResponse = evaluateDocumentJson(jsonData['Invoice'], 'InvoiceLine');
        break;
      case 'ar:ApplicationResponse':
        jsonResponse = evaluateApplicationResponseJson(jsonData['ar:ApplicationResponse']);
        break;
      case 'DebitNote':
        jsonResponse = evaluateDocumentJson(jsonData['DebitNote'], 'DebitNoteLine');
        break;
      default:
        console.log('Unknown JSON type:', identifier);
    }

    return jsonResponse;

  } catch (error) {
    console.error('Error extracting values:', error);
    return {
      invoiceNumber: 'Error',
      description: 'Error',
      rucEmi: 'Error',
      amountNoTax: 'Error',
      currencyCode: 'Error',
      issueDate: 'Error'
    };
  }
};

const evaluateDocumentJson = (documentJson, lineItemKey) => {
  const invoiceNumber = findValue(documentJson, equivalenceDict['cbc:ID']);
  const description = getDescription(documentJson, lineItemKey);
  const rucEmi = findValue(documentJson['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification'], equivalenceDict['cbc:ID']);
  const amountNoTax = findValue(documentJson['cac:TaxTotal']?.['cac:TaxSubtotal'], ['cbc:TaxableAmount']);
  const currencyCode = findValue(documentJson, equivalenceDict['cbc:DocumentCurrencyCode']);
  const issueDate = findValue(documentJson, equivalenceDict['cbc:IssueDate']);

  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax,
    currencyCode,
    issueDate
  };
};

const getDescription = (documentJson, lineItemKey) => {
  const lineItems = documentJson[`cac:${lineItemKey}`];
  if (Array.isArray(lineItems)) {
    return lineItems.map(lineItem => {
      const desc = lineItem['cac:Item']['cbc:Description'];
      return Array.isArray(desc) ? desc.join(' ') : desc;
    }).join(' ');
  } else {
    const desc = lineItems?.['cac:Item']?.['cbc:Description'];
    return Array.isArray(desc) ? desc.join(' ') : desc;
  }
};

const evaluateApplicationResponseJson = (appRespJson) => {
  const invoiceNumber = findValue(appRespJson['cac:DocumentResponse']?.['cac:DocumentReference'], equivalenceDict['cbc:ID']);
  const description = findValue(appRespJson['cac:DocumentResponse']?.['cac:Response'], equivalenceDict['cbc:Description']);
  const currencyCode = findValue(appRespJson, equivalenceDict['cbc:DocumentCurrencyCode']);
  const issueDate = findValue(appRespJson, equivalenceDict['cbc:IssueDate']);

  return {
    invoiceNumber,
    description,
    rucEmi: 'TBD',
    amountNoTax: 'TBD',
    currencyCode,
    issueDate
  };
};

// Function to read and convert all XML files in the folder, then write to Excel
const processXmlFiles = async () => {
  const workbook = xlsx.utils.book_new();
  const worksheetData = [['File', 'Invoice', 'Description', 'Ruc Emisor', 'Amount no tax', 'Currency Code', 'Issue Date']];

  try {
    const files = fs.readdirSync(xmlFolder);
    for (const file of files) {
      if (path.extname(file) === '.xml') {
        const filePath = path.join(xmlFolder, file);
        const xmlData = fs.readFileSync(filePath, 'utf8');
        const jsonData = await convertXmlToJson(xmlData);

        const values = extractValues(jsonData);
        // console.log(values);
        
        worksheetData.push([file, values.invoiceNumber, values.description, values.rucEmi, values.amountNoTax, values.currencyCode, values.issueDate]);

        console.log(`Processed file: ${file}`);
      }
    }

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, outputExcelFile);

    console.log(`Data has been written to ${outputExcelFile}`);
  } catch (error) {
    console.error('Error processing XML files:', error);
  }
};

// Execute the function
processXmlFiles();








// const fs = require('fs');
// const path = require('path');
// const xml2js = require('xml2js');
// const xlsx = require('xlsx');

// const xmlFolder = './XmlFiles'; // Folder containing XML files
// const outputExcelFile = './output.xlsx'; // Output Excel file

// // Equivalence dictionary for element names
// const equivalenceDict = {
//   'cbc:ID': ['cbc:ID', 'cbc:CompanyID', 'cbc:maku', 'cbc:demo'],
//   'cbc:Description': ['cbc:Description'],
//   'cbc:IssueDate': ['cbc:IssueDate'],
//   'cbc:DocumentCurrencyCode': ['cbc:DocumentCurrencyCode']
// };

// // Function to convert XML to JSON
// const convertXmlToJson = (xml) => {
//   return new Promise((resolve, reject) => {
//     xml2js.parseString(xml, { explicitArray: false }, (err, result) => {
//       if (err) {
//         reject(err);
//       } else {
//         resolve(result);
//       }
//     });
//   });
// };

// // Function to find the value of a key based on equivalence dictionary
// const findValue = (obj, keys) => {
//   for (const key of keys) {
//     if (obj[key]) {
//       return obj[key];
//     }
//   }
//   return 'Not found';
// };

// // Function to extract specific values from the JSON data
// const extractValues = (jsonData) => {
//   try {
//     const identifier = Object.keys(jsonData)[0] 
//     let jsonResponse = {};

//     switch (identifier) {
//         case 'Invoice':
//             jsonResponse = evaluateInvoiceJson(jsonData);
//             break;
//         case 'ar:ApplicationResponse':
//             jsonResponse = evaluateApplicationResponseJson(jsonData);
//             break;
//         case 'DebitNote':
//             jsonResponse = evaluateDebitNoteJson(jsonData);
//             break;        
//         default:
//             console.log('Unknown JSON type:', {});
//     }    

//     return jsonResponse;

//   } catch (error) {
//     console.error('Error extracting values:', error);
//     return {
//       description: 'Error',
//       id: 'Error',
//       issueDate: 'Error'
//     };
//   }
// };

// function evaluateInvoiceJson(invoiceJson){
//     let invoiceJsonReader = invoiceJson['Invoice'];

//     // let valueToDebug = invoiceJsonReader?.['cbc:ID'];
//     // console.log(valueToDebug);
    

//     let invoiceNumber = invoiceJsonReader?.['cbc:ID'];
//     let description = '';

//     if (Array.isArray(invoiceJsonReader?.['cac:InvoiceLine'])) {
//         description = invoiceJsonReader?.['cac:InvoiceLine'].map(invoiceLine => {
//             const description = invoiceLine['cac:Item']['cbc:Description'];
//             // If description is an array, join its elements into one string
//             if (Array.isArray(description)) {
//                 return description.join(' ');
//             } else {
//                 return description;
//             }
//         }).join(' ');        
//     } else {
//         description = invoiceJsonReader?.['cac:InvoiceLine']?.['cac:Item']?.['cbc:Description'];
//     }   
    

//     let rucEmi = invoiceJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
//     let amountNoTax = invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];    
//     let currencyCode = "";

//     // If currency code is a direct value
//     if (invoiceJsonReader.hasOwnProperty('DocumentCurrencyCode')) {
//         currencyCode = invoiceJsonReader.DocumentCurrencyCode;
//     }
//     // If currency code is nested within an object
//     else if (invoiceJsonReader.hasOwnProperty('cbc:DocumentCurrencyCode') && typeof invoiceJsonReader['cbc:DocumentCurrencyCode'] === 'object') {
//         currencyCode = invoiceJsonReader['cbc:DocumentCurrencyCode']['_'];
//     }
//     // If currency code is nested within an object with multiple attributes
//     else if (invoiceJsonReader.hasOwnProperty('cbc:DocumentCurrencyCode') && typeof invoiceJsonReader['cbc:DocumentCurrencyCode'] === 'string') {
//         currencyCode = invoiceJsonReader['cbc:DocumentCurrencyCode'];
//     }


//     let issueDate = findValue(invoiceJsonReader, equivalenceDict['cbc:IssueDate']);

//     return {
//         invoiceNumber,
//         description,
//         rucEmi,
//         amountNoTax,
//         currencyCode,
//         issueDate
//       };
// }

// function evaluateApplicationResponseJson(appRespJson){
//     let appRespJsonReader = appRespJson['ar:ApplicationResponse'];

//     let invoiceNumber = findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:DocumentReference'], equivalenceDict['cbc:ID']);
//     let description = findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:Response'], equivalenceDict['cbc:Description']);
//     let rucEmi = 'TBD';
//     let amountNoTax = 'TBD';
//     let currencyCode = findValue(appRespJsonReader, equivalenceDict['cbc:DocumentCurrencyCode']);
//     let issueDate = findValue(appRespJsonReader, equivalenceDict['cbc:IssueDate']);

//     return {
//         invoiceNumber,
//         description,
//         rucEmi,
//         amountNoTax,
//         currencyCode,
//         issueDate
//       };

// }

// function evaluateDebitNoteJson(debitNoteJson){

//     let debitNoteJsonReader = debitNoteJson['DebitNote'];


//     let invoiceNumber = debitNoteJsonReader?.['cbc:ID'];

//     let description = '';

//     if (Array.isArray(debitNoteJsonReader?.['cac:DebitNoteLine'])) {
//         description = debitNoteJsonReader?.['cac:DebitNoteLine'].map(invoiceLine => {
//             const description = invoiceLine['cac:Item']['cbc:Description'];
//             // If description is an array, join its elements into one string
//             if (Array.isArray(description)) {
//                 return description.join(' ');
//             } else {
//                 return description;
//             }
//         }).join(' ');        
//     } else {
//         let isArrayElement  = Array.isArray(debitNoteJsonReader?.['cac:DebitNoteLine']?.['cac:Item']?.['cbc:Description']);
//         if (isArrayElement)
//             description = debitNoteJsonReader?.['cac:DebitNoteLine']?.['cac:Item']?.['cbc:Description'].join(' ');
//         else
//             description = debitNoteJsonReader?.['cac:DebitNoteLine']?.['cac:Item']?.['cbc:Description'];
//     }

//     let rucEmi = debitNoteJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
//     let amountNoTax = debitNoteJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];

//     let currencyCode = "";

//     // If currency code is a direct value
//     if (debitNoteJsonReader.hasOwnProperty('DocumentCurrencyCode')) {
//         currencyCode = debitNoteJsonReader.DocumentCurrencyCode;
//     }
//     // If currency code is nested within an object
//     else if (debitNoteJsonReader.hasOwnProperty('cbc:DocumentCurrencyCode') && typeof debitNoteJsonReader['cbc:DocumentCurrencyCode'] === 'object') {
//         currencyCode = debitNoteJsonReader['cbc:DocumentCurrencyCode']['_'];
//     }
//     // If currency code is nested within an object with multiple attributes
//     else if (debitNoteJsonReader.hasOwnProperty('cbc:DocumentCurrencyCode') && typeof debitNoteJsonReader['cbc:DocumentCurrencyCode'] === 'string') {
//         currencyCode = debitNoteJsonReader['cbc:DocumentCurrencyCode'];
//     }    
    
//     let issueDate = findValue(debitNoteJsonReader, equivalenceDict['cbc:IssueDate']);
    
    

//     return {
//         invoiceNumber,
//         description,
//         rucEmi,
//         amountNoTax,
//         currencyCode,
//         issueDate
//       };

// }

// // Function to read and convert all XML files in the folder, then write to Excel
// const processXmlFiles = async () => {
//   const workbook = xlsx.utils.book_new();
//   const worksheetData = [['File', 'Invoice', 'Description', 'Ruc Emisor', 'Amount no tax', 'Currency Code', 'Issue Date']];

//   try {
//     const files = fs.readdirSync(xmlFolder);
//     for (const file of files) {
//       if (path.extname(file) === '.xml') {
//         const filePath = path.join(xmlFolder, file);
//         const xmlData = fs.readFileSync(filePath, 'utf8');
//         const jsonData = await convertXmlToJson(xmlData);

//         // console.log("%j", jsonData);
//         // console.log("----");
//         // console.log("--------");
//         // console.log("------------");
        
        

//         const values = extractValues(jsonData);
//         // console.log(values);
        
//         worksheetData.push([file, values.invoiceNumber, values.description, values.rucEmi, values.amountNoTax, values.currencyCode, values.issueDate]);

//         console.log(`Processed file: ${file}`);
//       }
//     }

//     const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
//     xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
//     xlsx.writeFile(workbook, outputExcelFile);

//     console.log(`Data has been written to ${outputExcelFile}`);
//   } catch (error) {
//     console.error('Error processing XML files:', error);
//   }
// };

// // Execute the function
// processXmlFiles();
