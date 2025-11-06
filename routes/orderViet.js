const express = require('express')
const { google } = require('googleapis')

const router = express.Router()

// Cấu hình credentials
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, '\n'),
  },
  scopes: ['https://www.googleapis.com/auth/spreadsheets'],
})

const sheets = google.sheets({ version: 'v4', auth })
const spreadsheetId = process.env.GOOGLE_SHEET_ID

// Helper function to generate bill code: ODVddmmyyhhmmss
function generateBillCode() {
  const date = new Date()
  const day = String(date.getDate()).padStart(2, '0')
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const year = String(date.getFullYear()).slice(-2)
  const hour = String(date.getHours()).padStart(2, '0')
  const minute = String(date.getMinutes()).padStart(2, '0')
  const second = String(date.getSeconds()).padStart(2, '0')

  return `ODV${day}${month}${year}${hour}${minute}${second}`
}

// Helper function to get sheet name
function getSheetName(month, year) {
  return `ORDVIET_${month}_${year}`
}

// Helper function to create sheet if not exists
async function createSheetIfNotExists(sheetName) {
  try {
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId,
    })

    const sheetExists = spreadsheet.data.sheets.some(
      (sheet) => sheet.properties.title === sheetName,
    )

    if (!sheetExists) {
      await sheets.spreadsheets.batchUpdate({
        spreadsheetId,
        resource: {
          requests: [
            {
              addSheet: {
                properties: {
                  title: sheetName,
                },
              },
            },
          ],
        },
      })

      // Add headers
      const headers = ['Mã bill', 'Hình ảnh bill', 'Status', 'Số lượng', 'Tổng thanh toán', 'Note']

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1:F1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [headers],
        },
      })
    }
  } catch (error) {
    console.error('Error creating sheet:', error)
    throw error
  }
}

// Get all bills for a month
router.get('/bills', async (req, res) => {
  try {
    const { month, year } = req.query

    if (!month || !year) {
      return res.status(400).json({
        success: false,
        message: 'Month and year are required',
      })
    }

    const sheetName = getSheetName(month, year)

    // Create sheet if not exists
    await createSheetIfNotExists(sheetName)

    // Get data
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:F`,
    })

    const rows = response.data.values || []
    const bills = rows
      .slice(1)
      .map((row, index) => ({
        rowIndex: index + 2, // +2 because header is row 1 and data starts from row 2
        billCode: row[0] || '',
        billImage: row[1] ? row[1].replace(/^=IMAGE\("(.+)"\)$/, '$1') : '',
        status: row[2] || '',
        quantity: parseInt(row[3]) || 0,
        totalAmount: parseInt(row[4]) || 0,
        note: row[5] || '',
        month: parseInt(month),
        year: parseInt(year),
      }))
      .filter((bill) => bill.billCode) // Filter out empty rows

    res.json({
      success: true,
      data: bills,
    })
  } catch (error) {
    console.error('Error getting bills:', error)
    res.status(500).json({
      success: false,
      message: error.message,
    })
  }
})

// Create a new bill
router.post('/bills', async (req, res) => {
  try {
    const { billImage, status, quantity, totalAmount, note, month, year } = req.body

    if (!month || !year) {
      return res.status(400).json({
        success: false,
        message: 'Month and year are required',
      })
    }

    const sheetName = getSheetName(month, year)

    // Create sheet if not exists
    await createSheetIfNotExists(sheetName)

    // Generate bill code
    const billCode = generateBillCode()

    // Get existing data to find next row
    const existingData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:F`,
    })

    const rows = existingData.data.values || []
    const nextRow = rows.length + 1

    // Prepare row data
    const rowData = [
      billCode,
      billImage ? `=IMAGE("${billImage}")` : '',
      status || 'ĐANG VẬN CHUYỂN',
      quantity || 0,
      totalAmount || 0,
      note || '',
    ]

    // Insert data
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${nextRow}:F${nextRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [rowData],
      },
    })

    res.json({
      success: true,
      data: { billCode },
      message: 'Bill created successfully',
    })
  } catch (error) {
    console.error('Error creating bill:', error)
    res.status(500).json({
      success: false,
      message: error.message,
    })
  }
})

// Update a bill
router.put('/bills/:billCode', async (req, res) => {
  try {
    const { billCode } = req.params
    const { billImage, status, quantity, totalAmount, note, month, year } = req.body

    if (!month || !year) {
      return res.status(400).json({
        success: false,
        message: 'Month and year are required',
      })
    }

    const sheetName = getSheetName(month, year)

    // Get existing data to find the bill
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:F`,
    })

    const rows = response.data.values || []
    const billRowIndex = rows.findIndex((row, index) => index > 0 && row[0] === billCode)

    if (billRowIndex === -1) {
      return res.status(404).json({
        success: false,
        message: 'Bill not found',
      })
    }

    const targetRow = billRowIndex + 1 // +1 because array is 0-indexed but sheet rows are 1-indexed

    // Get current data
    const currentRow = rows[billRowIndex]

    // Prepare updated data (keep existing values if not provided)
    const rowData = [
      billCode, // Keep bill code
      billImage !== undefined ? (billImage ? `=IMAGE("${billImage}")` : '') : currentRow[1],
      status !== undefined ? status : currentRow[2],
      quantity !== undefined ? quantity : currentRow[3],
      totalAmount !== undefined ? totalAmount : currentRow[4],
      note !== undefined ? note : currentRow[5],
    ]

    // Update data
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${targetRow}:F${targetRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [rowData],
      },
    })

    res.json({
      success: true,
      message: 'Bill updated successfully',
    })
  } catch (error) {
    console.error('Error updating bill:', error)
    res.status(500).json({
      success: false,
      message: error.message,
    })
  }
})

// Get orders with "HÀNG VIỆT" status from BÁN HÀNG and CTV sheets
router.get('/hang-viet-orders', async (req, res) => {
  try {
    const { months, year, customerType } = req.query

    if (!months || !year || !customerType) {
      return res.status(400).json({
        success: false,
        message: 'Months, year, and customerType are required',
      })
    }

    const monthArray = months.split(',').map((m) => parseInt(m))
    const sheetBaseName = customerType === 'customer' ? 'BÁN HÀNG' : 'CTV'

    let allOrders = []

    for (const month of monthArray) {
      const sheetName = `${sheetBaseName}_${month}_${year}`

      try {
        const response = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${sheetName}!A:O`,
        })

        const rows = response.data.values || []

        // Skip first 3 rows (headers) and process data
        const orders = rows
          .slice(3)
          .map((row, index) => ({
            rowIndex: index, // This is the actual row index used for updates
            date: row[0] || '',
            customerName: row[1] || '',
            productImage: row[2] ? row[2].replace(/^=IMAGE\("(.+)"\)$/, '$1') : '',
            productName: row[3] || '',
            color: row[4] || '',
            size: row[5] || '',
            quantity: row[6] || '',
            total: row[7] || '',
            status: row[8] || '',
            linkFb: row[9] || '',
            contactInfo: row[10] || '',
            note: row[11] || '',
            productCode: row[12] || '',
            orderCode: row[13] || '',
            shippingCode: row[14] || '',
            month,
            year: parseInt(year),
            sheetType: customerType,
          }))
          .filter((order) => order.status === 'HÀNG VIỆT' && order.customerName)

        allOrders = allOrders.concat(orders)
      } catch (error) {
        console.log(`Sheet ${sheetName} not found or error:`, error.message)
        // Continue with next month
      }
    }

    res.json({
      success: true,
      data: allOrders,
    })
  } catch (error) {
    console.error('Error getting hang viet orders:', error)
    res.status(500).json({
      success: false,
      message: error.message,
    })
  }
})

// Process orders: update status to "HÀNG VỀ" and add bill code
router.post('/process-orders', async (req, res) => {
  try {
    const { billCode, orderRowIndices, months, year, customerType } = req.body

    if (!billCode || !orderRowIndices || !months || !year || !customerType) {
      return res.status(400).json({
        success: false,
        message: 'All fields are required',
      })
    }

    const sheetBaseName = customerType === 'customer' ? 'BÁN HÀNG' : 'CTV'

    // Group orders by month
    const ordersByMonth = {}
    orderRowIndices.forEach((item) => {
      if (!ordersByMonth[item.month]) {
        ordersByMonth[item.month] = []
      }
      ordersByMonth[item.month].push(item.rowIndex)
    })

    // Update each order
    for (const [month, rowIndices] of Object.entries(ordersByMonth)) {
      const sheetName = `${sheetBaseName}_${month}_${year}`

      for (const rowIndex of rowIndices) {
        const targetRow = rowIndex + 4 // +4 because data starts from row 4

        // Update status (column I) and orderCode (column N)
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId,
          resource: {
            data: [
              {
                range: `${sheetName}!I${targetRow}`,
                values: [['HÀNG VỀ']],
              },
              {
                range: `${sheetName}!N${targetRow}`,
                values: [[billCode]],
              },
            ],
            valueInputOption: 'USER_ENTERED',
          },
        })
      }
    }

    // Update bill status in ORDVIET sheet
    const billMonth = months[0] // Assuming bill is in the first month
    const billSheetName = getSheetName(billMonth, year)

    // Find bill row
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${billSheetName}!A:F`,
    })

    const rows = response.data.values || []
    const billRowIndex = rows.findIndex((row, index) => index > 0 && row[0] === billCode)

    if (billRowIndex !== -1) {
      const targetRow = billRowIndex + 1

      // Update status to "HÀNG VỀ"
      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${billSheetName}!C${targetRow}`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [['HÀNG VỀ']],
        },
      })
    }

    res.json({
      success: true,
      message: 'Orders processed successfully',
    })
  } catch (error) {
    console.error('Error processing orders:', error)
    res.status(500).json({
      success: false,
      message: error.message,
    })
  }
})

module.exports = router
