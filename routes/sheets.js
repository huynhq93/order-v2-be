const express = require('express')
const { google } = require('googleapis')

const router = express.Router()

// Cấu hình credentials
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
});

const sheets = google.sheets({ version: 'v4', auth });
const spreadsheetId = process.env.GOOGLE_SHEET_ID;

router.get('/', async (req, res) => {
  try {
    let date = new Date()
    if (req.query.year && req.query.month) {
      date = new Date(Number.parseInt(req.query.year), Number.parseInt(req.query.month) - 1, 1)
    }
    const result = await readSheet(SHEET_TYPES[req.query.type], date)
    res.json({data: result})
  } catch (err) {
    console.error(err)
    res.status(500).send('Google Sheet Error')
  }
})

router.post('/', async (req, res) => {
  try {
    const {
      date,
      customerName,
      productImage,
      productName,
      color,
      size,
      quantity,
      total,
      status,
      linkFb,
      contactInfo,
      note,
    } = req.body;

    console.log("req.query.type:",req.query.type)
    const dateObj = date ? new Date(date) : new Date();

    const result = await readSheet(SHEET_TYPES[req.query.type], dateObj)
    let nextRow = result.length + 1
    nextRow = Math.max(nextRow, 4)

    const sheetName = getMonthlySheetName(SHEET_TYPES[req.query.type], dateObj);

    const range = `${sheetName}!A${nextRow}`;

    const values = [
      formatDateForSheet(dateObj),
      customerName,
      productImage ? `=IMAGE("${productImage}")` : '',
      productName,
      color,
      size,
      quantity,
      total,
      status,
      linkFb,
      contactInfo,
      note,
    ];

    const response = await appendSheet(range, values);
    res.json({ message: 'Order added successfully', data: response });
  } catch (error) {
    console.error('Lỗi khi thêm order:', error);
    res.status(500).json({ error: 'Failed to add order to Google Sheet' });
  }
});

function formatDateForSheet(date) {
  // Format dạng: dd/MM/yyyy hoặc ISO nếu bạn config Sheet đọc kiểu khác
  const d = new Date(date);
  return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`;
}

// Đọc dữ liệu từ sheet
async function readSheet(baseSheetName, date) {
  try {
    const sheetName = getMonthlySheetName(baseSheetName, date)

    const spreadsheetId = process.env.GOOGLE_SHEET_ID
    
    if (!spreadsheetId) {
      throw new Error("GOOGLE_SHEET_ID is not defined")
    }
    
    const start = new Date();

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: [`${sheetName}!A:Z`],
      includeGridData: true,
      fields: 'sheets.data.rowData.values.userEnteredValue',
    });
    const end = new Date();
    const durationMs = end.getTime() - start.getTime(); // thời gian chạy (ms)
    console.log(`Time taken get data gg: ${durationMs} ms`);
    
    const rows = response.data.sheets?.[0]?.data?.[0]?.rowData || [];
    
    // Map the data based on the sheet structure
    if (baseSheetName === SHEET_TYPES.ORDERS || baseSheetName === SHEET_TYPES.COLLABORATORS) {
      // Skip the first 3 rows (headers)
      return rows.slice(3).map((row, index) => {
        const cells = row.values || [];
        return {
          rowIndex: index,
          date: parseGoogleSheetDate(cells[0]),
          customerName: getCellString(cells[1]),
          productImage: extractImageUrl(getCellString(cells[2])),
          productName: getCellString(cells[3]),
          color: getCellString(cells[4]),
          size: getCellString(cells[5]),
          quantity: getCellString(cells[6]),
          total: getCellString(cells[7]),
          status: getCellString(cells[8]),
          linkFb: getCellString(cells[9]),
          contactInfo: getCellString(cells[10]),
          note: getCellString(cells[11]),
          month: `${date.getMonth() + 1}/${date.getFullYear()}`,
        }
      }).filter(item => item.customerName); // Filter out empty rows
    } else if (baseSheetName === SHEET_TYPES.INVENTORY) {
      // Skip the first row (header)
      return rows.slice(1).map((row, index) => {
        const cells = row.values || [];
        return {
          rowIndex: index,
          date: parseGoogleSheetDate(cells[0]),
          customerName: getCellString(cells[1]),
          productImage: extractImageUrl(getCellString(cells[2])),
          productName: getCellString(cells[3]),
          color: getCellString(cells[4]),
          size: getCellString(cells[5]),
          quantity: getCellString(cells[6]),
          total: getCellString(cells[7]),
          status: getCellString(cells[8]),
          linkFb: getCellString(cells[9]),
          contactInfo: getCellString(cells[10]),
          note: getCellString(cells[11]),
          month: `${date.getMonth() + 1}/${date.getFullYear()}`,
        }
      }).filter(item => item.productName); // Filter out empty rows
    }
    return []
  } catch (error) {
    console.error('Lỗi khi đọc dữ liệu:', error);
    throw error;
  }
}

// Ghi dữ liệu vào sheet
async function writeSheet(spreadsheetId, range, values) {
  try {
    const response = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      resource: { values },
    });
    return response.data;
  } catch (error) {
    console.error('Lỗi khi ghi dữ liệu:', error);
    throw error;
  }
}

// Thêm dữ liệu mới vào sheet
async function appendSheet(range, values) {
  try {
    const response = await sheets.spreadsheets.values.append({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      resource: { values: [values] },
    });
    return response.data;
  } catch (error) {
    console.error('Lỗi khi thêm dữ liệu:', error);
    throw error;
  }
}


function extractImageUrl(cellValue) {
  const match = cellValue.match(/^=IMAGE\("([^"]+)"\)/);
  
  return match ? match[1] : '';
}

function getCellString(cell) {
  const val = cell?.userEnteredValue;
  if (!val) return '';

  if (val.stringValue !== undefined) return val.stringValue;
  if (val.numberValue !== undefined) return String(val.numberValue);
  if (val.boolValue !== undefined) return String(val.boolValue);
  if (val.formulaValue !== undefined) return val.formulaValue;

  return '';
}

function parseGoogleSheetDate(cell) {
  const value = cell?.userEnteredValue;
  if (!value) return '';
  if (value.numberValue !== undefined) {
    // Google Sheets date serial number -> JS Date
    const baseDate = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
    const date = new Date(baseDate.getTime() + value.numberValue * 24 * 60 * 60 * 1000);
    return  `${date.getDate()}/${date.getMonth() + 1}`;
  } else if (value.stringValue !== undefined) {
    return value.stringValue; 
  } else {
    return ''; 
  }
}

function getMonthlySheetName(baseSheetName, date = new Date()) {
  const month = date.getMonth() + 1 // getMonth() returns 0-11
  const year = date.getFullYear()
  return `${baseSheetName}_${month}_${year}`
}

function formatMonthYear(date = new Date()) {
  return `${date.getMonth() + 1}/${date.getFullYear()}`
}

const SHEET_TYPES = {
  ORDERS: "BÁN HÀNG",
  INVENTORY: "NHẬP HÀNG",
  COLLABORATORS: "CÔNG TÁC VIÊN",
}

module.exports = router

// export const googleSheetsService = {
//   // Đọc dữ liệu từ sheet
//   async readSheet(spreadsheetId, range) {
//     try {
//       const sheetName = getMonthlySheetName(baseSheetName, date)
//       // In a real implementation, this would be:
//       const sheets = await getGoogleSheetsClient()
//       const spreadsheetId = process.env.GOOGLE_SHEET_ID
      
//       if (!spreadsheetId) {
//         throw new Error("GOOGLE_SHEET_ID is not defined")
//       }
      
//       const start = new Date();

//       const response = await sheets.spreadsheets.get({
//         spreadsheetId,
//         ranges: [`${sheetName}!A:Z`],
//         includeGridData: true,
//         fields: 'sheets.data.rowData.values.userEnteredValue',
//       });
//       const end = new Date();
//       const durationMs = end.getTime() - start.getTime(); // thời gian chạy (ms)
//       console.log(`Time taken get data gg: ${durationMs} ms`);
      
//       const rows = response.data.sheets?.[0]?.data?.[0]?.rowData || [];
      
//       // Map the data based on the sheet structure
//       if (baseSheetName === SHEET_TYPES.ORDERS || baseSheetName === SHEET_TYPES.COLLABORATORS) {
//         // Skip the first 3 rows (headers)
//         return rows.slice(3).map((row, index) => {
//           const cells = row.values || [];
//           return {
//             rowIndex: index,
//             date: parseGoogleSheetDate(cells[0]),
//             customerName: getCellString(cells[1]),
//             productImage: extractImageUrl(getCellString(cells[2])),
//             productName: getCellString(cells[3]),
//             color: getCellString(cells[4]),
//             size: getCellString(cells[5]),
//             quantity: getCellString(cells[6]),
//             total: getCellString(cells[7]),
//             status: getCellString(cells[8]),
//             linkFb: getCellString(cells[9]),
//             contactInfo: getCellString(cells[10]),
//             note: getCellString(cells[11]),
//             month: `${date.getMonth() + 1}/${date.getFullYear()}`,
//           }
//         }).filter(item => item.customerName); // Filter out empty rows
//       } else if (baseSheetName === SHEET_TYPES.INVENTORY) {
//         // Skip the first row (header)
//         return rows.slice(1).map((row, index) => {
//           const cells = row.values || [];
//           return {
//             rowIndex: index,
//             date: parseGoogleSheetDate(cells[0]),
//             customerName: getCellString(cells[1]),
//             productImage: extractImageUrl(getCellString(cells[2])),
//             productName: getCellString(cells[3]),
//             color: getCellString(cells[4]),
//             size: getCellString(cells[5]),
//             quantity: getCellString(cells[6]),
//             total: getCellString(cells[7]),
//             status: getCellString(cells[8]),
//             linkFb: getCellString(cells[9]),
//             contactInfo: getCellString(cells[10]),
//             note: getCellString(cells[11]),
//             month: `${date.getMonth() + 1}/${date.getFullYear()}`,
//           }
//         }).filter(item => item.productName); // Filter out empty rows
//       }
//       return []
//     } catch (error) {
//       console.error('Lỗi khi đọc dữ liệu:', error);
//       throw error;
//     }
//   },

//   // Ghi dữ liệu vào sheet
//   async writeSheet(spreadsheetId, range, values) {
//     try {
//       const response = await sheets.spreadsheets.values.update({
//         spreadsheetId,
//         range,
//         valueInputOption: 'RAW',
//         resource: { values },
//       });
//       return response.data;
//     } catch (error) {
//       console.error('Lỗi khi ghi dữ liệu:', error);
//       throw error;
//     }
//   },

//   // Thêm dữ liệu mới vào sheet
//   async appendSheet(spreadsheetId, range, values) {
//     try {
//       const response = await sheets.spreadsheets.values.append({
//         spreadsheetId,
//         range,
//         valueInputOption: 'RAW',
//         resource: { values },
//       });
//       return response.data;
//     } catch (error) {
//       console.error('Lỗi khi thêm dữ liệu:', error);
//       throw error;
//     }
//   },


//   extractImageUrl(cellValue) {
//     const match = cellValue.match(/^=IMAGE\("([^"]+)"\)/);
    
//     return match ? match[1] : '';
//   },

//   getCellString(cell) {
//     const val = cell?.userEnteredValue;
//     if (!val) return '';

//     if (val.stringValue !== undefined) return val.stringValue;
//     if (val.numberValue !== undefined) return String(val.numberValue);
//     if (val.boolValue !== undefined) return String(val.boolValue);
//     if (val.formulaValue !== undefined) return val.formulaValue;

//     return '';
//   },

//   parseGoogleSheetDate(cell) {
//     const value = cell?.userEnteredValue;
//     if (!value) return '';
//     if (value.numberValue !== undefined) {
//       // Google Sheets date serial number -> JS Date
//       const baseDate = new Date(Date.UTC(1899, 11, 30)); // 1899-12-30
//       const date = new Date(baseDate.getTime() + value.numberValue * 24 * 60 * 60 * 1000);
//       return  `${date.getDate()}/${date.getMonth() + 1}`;
//     } else if (value.stringValue !== undefined) {
//       return value.stringValue; 
//     } else {
//       return ''; 
//     }
//   },

//   getMonthlySheetName(baseSheetName, date = new Date()) {
//     const month = date.getMonth() + 1 // getMonth() returns 0-11
//     const year = date.getFullYear()
//     return `${baseSheetName}_${month}_${year}`
//   },
  
//   formatMonthYear(date = new Date()) {
//     return `${date.getMonth() + 1}/${date.getFullYear()}`
//   }
// }; 