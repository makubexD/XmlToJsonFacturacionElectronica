const fs = require('fs');
const path = require('path');
const xml2js = require('xml2js');
const xlsx = require('xlsx');
const pdf = require('pdf-parse');
const axios = require('axios');

// Generic directory containing both XML and PDF files
const genericFolder = './Files'; // Modify this path as needed
const outputExcelFile = './output.xlsx'; // Output Excel file

const delay = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// Equivalence dictionary for XML element names
const equivalenceDict = {
  'cbc:ID': ['cbc:ID', 'cbc:CompanyID', 'cbc:maku', 'cbc:demo'],
  'cbc:Description': ['cbc:Description'],
  'cbc:IssueDate': ['cbc:IssueDate'],
  'cbc:DocumentCurrencyCode': ['cbc:DocumentCurrencyCode']
};

class TokenManager {
  constructor(tokens) {
    this.tokens = tokens;
    this.currentIndex = 0;
  }

  getNextToken() {
    const token = this.tokens[this.currentIndex];
    this.currentIndex = (this.currentIndex + 1) % this.tokens.length;
    return token;
  }
}

class InvoiceData {
  constructor(type, filePath) {
    this.type = type;
    this.filePath = filePath;
    this.invoiceNumber = 'N/A';
    this.description = 'N/A';
    this.rucEmi = 'N/A';
    this.amountNoTax = 'N/A';
    this.currencyCode = 'N/A';
    this.issueDate = 'N/A';
    this.razonSocial = 'N/A';
    this.inquilino = 'N/A';
    this.montoAlquiler = 'N/A';
    this.tributoResultante = 'N/A';
    this.fechaPago = 'N/A';
    this.StackError = '';
  }

  addError(error) {
    this.StackError += `${error.message} in ${this.filePath} `;
  }

  populateData(data) {
    Object.assign(this, data);
  }
}

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

// Function to find the value of a key based on the equivalence dictionary
const findValue = (obj, keys) => {
  for (const key of keys) {
    if (obj[key]) {
      return typeof obj[key] === 'object' && obj[key]['_'] ? obj[key]['_'] : obj[key];
    }
  }
  return 'Not found';
};


// function findValueByPath(json, path) {
//   const keys = path.split('.'); // Split the path into keys
//   let current = json; // Start from the root of the JSON

//   for (const key of keys) {
//       // If current is an array, we need to search through its items
//       if (Array.isArray(current)) {
//           // Create an array to hold found results
//           let results = [];
//           for (const item of current) {
//               // Recursively search in each item
//               const result = findValueByPath(item, key);
//               if (result !== undefined) {
//                   results.push(result); // Collect found values
//               }
//           }
//           if (results.length > 0) {
//               current = results; // If we found results, continue with them
//           } else {
//               return undefined; // No values found
//           }
//       } else if (typeof current === 'object' && current !== null && key in current) {
//           // If the current level is an object, move deeper
//           current = current[key];
//       } else {
//           // Key not found, return undefined
//           return undefined;
//       }
//   }

//   // After navigating through the path, return the final value
//   // If the final value is an array, return its first element
//   if (Array.isArray(current)) {
//       return current[0]; // Return the first element of the array
//   }

//   // If it's an object with an underscore property, return that value
//   if (typeof current === 'object' && current !== null && '_' in current) {
//       return current._; // Return the value of the underscore property
//   }

//   // Return the final value
//   return current;
// }

/*

v2 working*/

function findValueByPath(json, path) {
  const keys = path.split('.'); // Split the path into keys
  let current = json; // Start from the root of the JSON

  for (const key of keys) {
      // If current is an array, we need to search through its items
      if (Array.isArray(current)) {
          let results = [];
          for (const item of current) {
              // Recursively search in each item
              const result = findValueByPath(item, key);
              if (result !== undefined) {
                  results.push(result); // Collect found values
              }
          }
          if (results.length > 0) {
              current = results; // If we found results, continue with them
          } else {
              return undefined; // No values found
          }
      } else if (typeof current === 'object' && current !== null) {
          // Check for the key directly
          if (key in current) {
              current = current[key];
          } 
          // If not found, check for the key with an underscore
          else if (`_${key}` in current) {
              current = current[`_${key}`];
          } else {
              // Key not found, return undefined
              return undefined;
          }
      } else {
          // Not an object or array, return undefined
          return undefined;
      }
  }

  // After navigating through the path, handle the final value
  if (Array.isArray(current)) {
      if (current.length === 0) return undefined; // Return undefined if empty array
      const firstElement = current[0];
      // Check if the first element is an object with an underscore property
      if (typeof firstElement === 'object' && firstElement !== null) {
          return firstElement._ !== undefined ? firstElement._ : firstElement; // Return underscore property or the object itself
      }
      return firstElement; // Return the first element directly
  }

  // If current is an object, check for an underscore property
  if (typeof current === 'object' && current !== null && '_' in current) {
      return current._; // Return the value of the underscore property
  }

  // Return the final value directly if it's not an array or object with underscore
  return current;
}


/*
Version Working final description concat
*/



//function findValueByPathLastNodeMerge(json, path, mergeLastNode = false) {
  function getValueFromPath(json, path) {
    // Split the path by "." to create an array of keys
    const keys = path.split('.');
  
    let current = json;
  
    // Traverse the JSON object based on the keys array
    for (let key of keys) {
      if (Array.isArray(current)) {
        // If it's an array, map and extract values for the given key
        current = current.map(item => item[key]).flat().filter(Boolean);
      } else {
        current = current[key] || [];
      }
    }
  
    // If the final value is still an array, join the string elements
    return Array.isArray(current)
      ? current.map(value => (typeof value === 'string' ? value : '')).join(' ')
      : current;
  }
  
  
  
  







const processXmlFile = async (filePath, tokenManager) => {
  const fileName = path.basename(filePath);
  const invoiceData = new InvoiceData('XML', fileName);

  try {
    const xmlData = fs.readFileSync(filePath, 'utf8');
    const parser = new xml2js.Parser();
    // const jsonData = await convertXmlToJson(xmlData);
    const jsonData = await parser.parseStringPromise(xmlData); // This will throw an error if the XML is invalid

    const values = extractValues(jsonData, invoiceData);  // Pass invoiceData for error handling

    if (values.rucEmi && values.rucEmi.length == 11){
      const apiResponse = await apiRequest(values.rucEmi, tokenManager);
      values.razonSocial = apiResponse.razonSocial;
    }   

    invoiceData.populateData(values);  // Populate data if no error
    console.log(`Processed XML file: ${fileName}`);
  } catch (error) {
    invoiceData.addError(error);  // Error handling here
    console.error(`Error processing XML file ${filePath}:`, error);
  }

  return invoiceData;
};

const extractValues = (jsonData, invoiceData) => {
  try {
    const identifier = Object.keys(jsonData)[0];
    let values = {};

    switch (identifier) {
      case 'Invoice':
        values = evaluateInvoiceJson(jsonData, identifier);
        break;
      case 'ar:ApplicationResponse':
        values = evaluateArApplicationResponseJson(jsonData, identifier);
        break;
      case 'ApplicationResponse':
        values = evaluateApplicationResponseJson(jsonData, identifier);
        break;
      case 'DebitNote':
        values = evaluateDebitNoteJson(jsonData, identifier);
        break;
      case 'PPLDocument':
        values = evaluatePPLDocumentJson(jsonData, identifier);
        break;
      default:
        throw new Error(`The XML has not defined structure : ${identifier}`);
    }    

    return values;
  } catch (error) {
    invoiceData.addError(error);  // Handle error properly in the calling function
    console.error('Error extracting values:', error);
    return {};  // Return an empty object in case of error
  }
};



function evaluateInvoiceJson(invoiceJson, identifier) {
  let invoiceJsonReader = invoiceJson[identifier];

  // console.dir(invoiceJsonReader, {depth:null});
  

  let invoiceNumber = findValueByPath(invoiceJsonReader, 'cbc:ID');//invoiceJsonReader?.['cbc:ID'];    
  let description = getValueFromPath(invoiceJsonReader, 'cac:InvoiceLine.cac:Item.cbc:Description');  
  let rucEmi = findValueByPath(invoiceJsonReader, 'cac:AccountingSupplierParty.cac:Party.cac:PartyIdentification.cbc:ID'); //invoiceJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
  let amountNoTax = findValueByPath(invoiceJsonReader, 'cac:TaxTotal.cac:TaxSubtotal.cbc:TaxableAmount'); //invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];
  let currencyCode = findValueByPath(invoiceJsonReader, 'cbc:DocumentCurrencyCode'); //invoiceJsonReader['cbc:DocumentCurrencyCode']?.['_'] || invoiceJsonReader['cbc:DocumentCurrencyCode'] || '';
  let issueDate = findValueByPath(invoiceJsonReader, 'cbc:IssueDate') //findValue(invoiceJsonReader, equivalenceDict['cbc:IssueDate']);

  
  // let description = ''
  // if (Array.isArray(invoiceJsonReader?.['cac:InvoiceLine'])) {
  //   description = invoiceJsonReader?.['cac:InvoiceLine'].map(invoiceLine => {
  //     const description = invoiceLine['cac:Item']['cbc:Description'];
  //     return Array.isArray(description) ? description.join(' ') : description;
  //   }).join(' ');
  // } else {
  //   description = invoiceJsonReader?.['cac:InvoiceLine']?.['cac:Item']?.['cbc:Description'];
  // }

  

 
  
  // console.log(invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']);

  // let amountNoTax = '';
  // let taxSubtotalMain = invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal'];
  // if (Array.isArray(taxSubtotalMain)) {    
  //   amountNoTax = taxSubtotalMain.map(item => {
  //     const taxableAmount = item?.['cbc:TaxableAmount']?.["_"];
  //     if (taxableAmount) {
  //       return taxableAmount;
  //     }
  //     return null;
  //   }).filter(amount => amount !== null);
  // } else if (taxSubtotalMain) {    
  //   const amountNoTaxValid = taxSubtotalMain?.['cbc:TaxableAmount']?.["_"];
  //   amountNoTax = amountNoTaxValid ? amountNoTaxValid : null; 
  // }

  
  // let taxSubtotalMain = invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal'];

  // if (Array.isArray(taxSubtotalMain)) {
  //   // Use reduce to return the first valid float and stop the reduce
  //   amountNoTax = taxSubtotalMain.reduce((acc, item) => {
  //     if (acc) return acc; // If already found, skip
  //     const taxableAmount = item?.['cbc:TaxableAmount']?.["_"];
  //     return taxableAmount && !isNaN(parseFloat(taxableAmount)) ? taxableAmount : null;
  //   }, null);
  // } else if (taxSubtotalMain) {    
  //   const amountNoTaxValid = taxSubtotalMain?.['cbc:TaxableAmount']?.["_"];
  //   amountNoTax = amountNoTaxValid && !isNaN(parseFloat(amountNoTaxValid)) ? amountNoTaxValid : null;
  // }
  
  
  // let amountNoTax = invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];
  
  

  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax : (amountNoTax || '').toString(),
    currencyCode,
    issueDate,
    StackError: ''
  };
}

function evaluateArApplicationResponseJson(appRespJson, identifier) {
  let appRespJsonReader = appRespJson[identifier];

  let invoiceNumber = findValueByPath(appRespJsonReader, 'cac:DocumentResponse.cac:DocumentReference.cbc:ID');//findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:DocumentReference'], equivalenceDict['cbc:ID']);
  let description = findValueByPath(appRespJsonReader, 'cac:DocumentResponse.cac:Response.cbc:Description');//findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:Response'], equivalenceDict['cbc:Description']);
  let rucEmi = 'TBD';
  let amountNoTax = 'TBD';
  let currencyCode = findValueByPath(appRespJsonReader, 'cbc:DocumentCurrencyCode');//findValue(appRespJsonReader, equivalenceDict['cbc:DocumentCurrencyCode']);
  let issueDate = findValueByPath(appRespJsonReader, 'cbc:IssueDate');//findValue(appRespJsonReader, equivalenceDict['cbc:IssueDate']);

  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax,
    currencyCode,
    issueDate,
    StackError: ''
  };
}

function evaluateApplicationResponseJson(appRespJson, identifier) {
  let appRespJsonReader = appRespJson[identifier];

  let invoiceNumber = findValueByPath(appRespJsonReader, 'cac:DocumentResponse.cac:DocumentReference.cbc:ID');//findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:DocumentReference'], equivalenceDict['cbc:ID']);
  let description = findValueByPath(appRespJsonReader, 'cac:DocumentResponse.cac:Response.cbc:Description');//findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:Response'], equivalenceDict['cbc:Description']);
  let rucEmi = 'TBD';
  let amountNoTax = 'TBD';
  let currencyCode = findValueByPath(appRespJsonReader, 'cbc:DocumentCurrencyCode');//findValue(appRespJsonReader, equivalenceDict['cbc:DocumentCurrencyCode']);
  let issueDate = findValueByPath(appRespJsonReader, 'cbc:IssueDate');//findValue(appRespJsonReader, equivalenceDict['cbc:IssueDate']);

  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax,
    currencyCode,
    issueDate,
    StackError: ''
  };
}

function evaluateDebitNoteJson(debitNoteJson, identifier) {
  let debitNoteJsonReader = debitNoteJson[identifier];

  let invoiceNumber = findValueByPath(debitNoteJsonReader, 'cbc:ID'); //debitNoteJsonReader?.['cbc:ID'];
  let description = findValueByPath(invoiceJsonReader, 'cac:DebitNoteLine.cac:Item.cbc:Description');
  

  // if (Array.isArray(debitNoteJsonReader?.['cac:DebitNoteLine'])) {
  //   description = debitNoteJsonReader?.['cac:DebitNoteLine'].map(invoiceLine => {
  //     const description = invoiceLine['cac:Item']['cbc:Description'];
  //     return Array.isArray(description) ? description.join(' ') : description;
  //   }).join(' ');
  // } else {
  //   description = debitNoteJsonReader?.['cac:DebitNoteLine']?.['cac:Item']?.['cbc:Description'];
  // }

  // let rucEmi = debitNoteJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
  // let amountNoTax = debitNoteJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];
  // let currencyCode = debitNoteJsonReader['cbc:DocumentCurrencyCode']?.['_'] || debitNoteJsonReader['cbc:DocumentCurrencyCode'] || '';
  // let issueDate = findValue(debitNoteJsonReader, equivalenceDict['cbc:IssueDate']);

  let rucEmi = findValueByPath(invoiceJsonReader,'cac:AccountingSupplierParty.cac:Party.cac:PartyIdentification.cbc:ID')
  let amountNoTax = findValueByPath(invoiceJsonReader,'cac:TaxTotal.cac:TaxSubtotal.cbc:TaxableAmount')
  let currencyCode = findValueByPath(invoiceJsonReader,'cbc:DocumentCurrencyCode')
  let issueDate = findValueByPath(invoiceJsonReader,'cbc:IssueDate')



  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax,
    currencyCode,
    issueDate,
    StackError: ''
  };
}


function evaluatePPLDocumentJson(pplDocumentJson, identifier) {
  let debitNoteJsonReader = pplDocumentJson[identifier];

  // console.dir(debitNoteJsonReader, {depth:null});

  const mainPath = 'ClientDocument.Invoice.';

  let invoiceNumber = findValueByPath(debitNoteJsonReader, `${mainPath}cbc:ID`);
  let description = findValueByPath(debitNoteJsonReader, `${mainPath}cac:InvoiceLine.cac:Item.cbc:Description`);
  let rucEmi = findValueByPath(debitNoteJsonReader, `${mainPath}cac:Signature.cac:SignatoryParty.cac:PartyIdentification.cbc:ID`);
  let amountNoTax = findValueByPath(debitNoteJsonReader, `${mainPath}cac:TaxTotal.cac:TaxSubtotal.cbc:TaxableAmount`);
  let currencyCode = findValueByPath(debitNoteJsonReader, `${mainPath}cbc:DocumentCurrencyCode`);
  let issueDate = findValueByPath(debitNoteJsonReader, `${mainPath}cbc:IssueDate`);


  return {
    invoiceNumber,
    description,
    rucEmi,
    amountNoTax,
    currencyCode,
    issueDate,
    StackError: ''
  };
}


const apiRequest = async (ruc, tokenManager) => {
  const token = tokenManager.getNextToken();
  try {
    const response = await axios({
      method: 'get',
      maxBodyLength: Infinity,
      url: `https://api.apis.net.pe/v2/sunat/ruc/full?numero=${ruc}`, 
      headers: {
        'Authorization': token
      }
    });
    // console.log(token);
    
    return response.data;
  } catch (err) {
    if (err.response && err.response.status === 429) {
      console.error('Rate limit hit (429), retrying after delay...');
      await delay(1000); // Delay before retrying
      return apiRequest(ruc, tokenManager); // Retry after delay
    }
    throw err; // Re-throw error if it's not a rate limit issue
  }
};


const processPdfFile = async (filePath, tokenManager) => {
  const fileName = path.basename(filePath);
  const invoiceData = new InvoiceData('PDF', fileName);

  try {
    const dataBuffer = fs.readFileSync(filePath);
    const data = await pdf(dataBuffer);
    const text = data.text;

    // Regexes to extract data from PDF
    const rucRegex = /RUC:\s*([\d\s\n]{1,12})/g;
    const inquilinoRegex = /(?:^|\n)Inquilino:\s*([^\n]+)/g;
    const montoAlquilerRegex = /Monto de Alquiler:\s*S\/\s*([\d,]+.\d+)/g;
    const tributoResultanteRegex = /Tributo Resultante:\s*S\/\s*([\d,]+.\d+)/g;
    // const fechaPagoRegex = /Fecha de Pago:\s*([^\n]+)/g;
    const fechaPagoRegex = /Fecha de Pago:\s*(\d{2}\/\d{2}\/\d{4})/g;;

    const rucMatch = rucRegex.exec(text);
    const inquilinoMatch = inquilinoRegex.exec(text);
    const montoAlquilerMatch = montoAlquilerRegex.exec(text);
    const tributoResultanteMatch = tributoResultanteRegex.exec(text);
    const fechaPagoMatch = fechaPagoRegex.exec(text);

    if (rucMatch) {
      const ruc = rucMatch[1].replace(/\s+/g, '').trim();
      if (ruc.length === 11) {
        const apiResponse = await apiRequest(ruc, tokenManager);
        invoiceData.populateData({
          rucEmi: ruc,
          razonSocial: apiResponse.razonSocial,
          inquilino: inquilinoMatch ? inquilinoMatch[1] : 'N/A',
          montoAlquiler: montoAlquilerMatch ? montoAlquilerMatch[1] : 'N/A',
          tributoResultante: tributoResultanteMatch ? tributoResultanteMatch[1] : 'N/A',
          fechaPago: fechaPagoMatch ? fechaPagoMatch[1] : 'N/A'
        });
      } else {
        invoiceData.addError(new Error(`Invalid RUC format in ${fileName}`));
      }
    } else {
      invoiceData.addError(new Error(`No RUC found in ${fileName}`));
    }
  } catch (error) {
    invoiceData.addError(error);
  }

  return invoiceData;
};

const processAllFiles = async () => {
  const workbook = xlsx.utils.book_new();
  const worksheetData = [['Type', 'File', 'Invoice', 'Description', 'Ruc Emisor', 'Amount no tax', 'Currency Code', 'Issue Date', 'Razon Social', 'Inquilino', 'Monto de Alquiler', 'Tributo Resultante', 'Fecha de Pago', 'StackError']];

  const tokenManager = new TokenManager([
    ''
  ]);

  try {
    const files = fs.readdirSync(genericFolder);
    for (const file of files) {
      const filePath = path.join(genericFolder, file);
      const ext = path.extname(file).toLowerCase();
      let result;

      if (ext === '.xml') {
        result = await processXmlFile(filePath, tokenManager);
      } else if (ext === '.pdf') {
        result = await processPdfFile(filePath, tokenManager);
      }

      if (result) {
        worksheetData.push([
          result.type,
          result.filePath,
          result.invoiceNumber,
          result.description,
          result.rucEmi,
          result.amountNoTax,
          result.currencyCode,
          result.issueDate,
          result.razonSocial,
          result.inquilino,
          result.montoAlquiler,
          result.tributoResultante,
          result.fechaPago,
          result.StackError
        ]);
      }
    }

    const worksheet = xlsx.utils.aoa_to_sheet(worksheetData);
    xlsx.utils.book_append_sheet(workbook, worksheet, 'Sheet1');
    xlsx.writeFile(workbook, outputExcelFile);
    console.log(`Data has been written to ${outputExcelFile}`);
  } catch (error) {
    console.error('Error processing files:', error);
  }
};

processAllFiles();

