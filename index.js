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

const processXmlFile = async (filePath, tokenManager) => {
  const fileName = path.basename(filePath);
  const invoiceData = new InvoiceData('XML', fileName);

  try {
    const xmlData = fs.readFileSync(filePath, 'utf8');
    const jsonData = await convertXmlToJson(xmlData);
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
        values = evaluateInvoiceJson(jsonData);
        break;
      case 'ar:ApplicationResponse':
        values = evaluateApplicationResponseJson(jsonData);
        break;
      case 'DebitNote':
        values = evaluateDebitNoteJson(jsonData);
        break;
      default:
        throw new Error(`Unknown JSON identifier: ${identifier}`);
    }    

    return values;
  } catch (error) {
    invoiceData.addError(error);  // Handle error properly in the calling function
    console.error('Error extracting values:', error);
    return {};  // Return an empty object in case of error
  }
};



function evaluateInvoiceJson(invoiceJson) {
  let invoiceJsonReader = invoiceJson['Invoice'];

  let invoiceNumber = invoiceJsonReader?.['cbc:ID'];
  let description = '';

  if (Array.isArray(invoiceJsonReader?.['cac:InvoiceLine'])) {
    description = invoiceJsonReader?.['cac:InvoiceLine'].map(invoiceLine => {
      const description = invoiceLine['cac:Item']['cbc:Description'];
      return Array.isArray(description) ? description.join(' ') : description;
    }).join(' ');
  } else {
    description = invoiceJsonReader?.['cac:InvoiceLine']?.['cac:Item']?.['cbc:Description'];
  }

  let rucEmi = invoiceJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
  let amountNoTax = invoiceJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];
  let currencyCode = invoiceJsonReader['cbc:DocumentCurrencyCode']?.['_'] || invoiceJsonReader['cbc:DocumentCurrencyCode'] || '';
  let issueDate = findValue(invoiceJsonReader, equivalenceDict['cbc:IssueDate']);

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

function evaluateApplicationResponseJson(appRespJson) {
  let appRespJsonReader = appRespJson['ar:ApplicationResponse'];

  let invoiceNumber = findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:DocumentReference'], equivalenceDict['cbc:ID']);
  let description = findValue(appRespJsonReader?.['cac:DocumentResponse']?.['cac:Response'], equivalenceDict['cbc:Description']);
  let rucEmi = 'TBD';
  let amountNoTax = 'TBD';
  let currencyCode = findValue(appRespJsonReader, equivalenceDict['cbc:DocumentCurrencyCode']);
  let issueDate = findValue(appRespJsonReader, equivalenceDict['cbc:IssueDate']);

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

function evaluateDebitNoteJson(debitNoteJson) {
  let debitNoteJsonReader = debitNoteJson['DebitNote'];

  let invoiceNumber = debitNoteJsonReader?.['cbc:ID'];
  let description = '';

  if (Array.isArray(debitNoteJsonReader?.['cac:DebitNoteLine'])) {
    description = debitNoteJsonReader?.['cac:DebitNoteLine'].map(invoiceLine => {
      const description = invoiceLine['cac:Item']['cbc:Description'];
      return Array.isArray(description) ? description.join(' ') : description;
    }).join(' ');
  } else {
    description = debitNoteJsonReader?.['cac:DebitNoteLine']?.['cac:Item']?.['cbc:Description'];
  }

  let rucEmi = debitNoteJsonReader?.['cac:AccountingSupplierParty']?.['cac:Party']?.['cac:PartyIdentification']?.['cbc:ID']["_"];
  let amountNoTax = debitNoteJsonReader?.['cac:TaxTotal']?.['cac:TaxSubtotal']?.['cbc:TaxableAmount']["_"];
  let currencyCode = debitNoteJsonReader['cbc:DocumentCurrencyCode']?.['_'] || debitNoteJsonReader['cbc:DocumentCurrencyCode'] || '';
  let issueDate = findValue(debitNoteJsonReader, equivalenceDict['cbc:IssueDate']);

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

