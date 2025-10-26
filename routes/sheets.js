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

// Helper function to generate product code: SP{year}{month}{date}{hour}{minute}{second}
function generateProductCode() {
  const date = new Date()
  const year = date.getFullYear()
  const month = String(date.getMonth() + 1).padStart(2, '0')
  const day = String(date.getDate()).padStart(2, '0')
  const hour = String(date.getHours()).padStart(2, '0')
  const minute = String(date.getMinutes()).padStart(2, '0')
  const second = String(date.getSeconds()).padStart(2, '0')
  
  return `SP${year}${month}${day}${hour}${minute}${second}`
}

router.get('/', async (req, res) => {
  try {
    let date = new Date()
    if (req.query.year && req.query.month) {
      date = new Date(Number.parseInt(req.query.year), Number.parseInt(req.query.month) - 1, 1)
    }

    const result = await readSheet(SHEET_TYPES[req.query.type], date)

    res.json({ data: result })
  } catch (err) {
    console.error('Detailed error:', err)
    res.status(500).json({
      error: 'Google Sheet Error',
      message: err.message,
      stack: err.stack,
    })
  }
})

router.post('/', async (req, res) => {
  try {
    const {
      date,
      customerName,
      productCode,
      orderCode,
      shippingCode,
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
    } = req.body

    console.log('req.query.type:', req.query.type)
    const dateObj = date ? parseDateString(date) : new Date()
    let generatedProductCode = '' // Declare variable at function scope

    // Logic add sản phẩm mới: chỉ khi KHÔNG có productCode và có productImage + productName
    if (!!(!productCode && productImage)) {
      try {
        // Generate unique product code
        generatedProductCode = generateProductCode() // Use new format: SP{year}{month}{date}{hour}{minute}{second}

        // Add to products sheet
        const existingProducts = await readSheet(SHEET_TYPES.PRODUCTS, dateObj)
        const productNextRow = existingProducts.length + 2
        const productSheetName = getMonthlySheetName(SHEET_TYPES.PRODUCTS, dateObj)
        const productRange = `${productSheetName}!A${productNextRow}`
        const productValues = [
          formatDateForSheet(dateObj),
          generatedProductCode,
          productImage ? `=IMAGE("${productImage}")` : '',
          productName,
        ]

        await appendSheet(productRange, productValues)
        console.log(`Added new product to sheet: ${generatedProductCode}`)
      } catch (error) {
        console.error('Failed to add product to sheet:', error)
        res.status(500).json({ error: 'Failed to add product to Google Sheet' })
        return
        // Continue with order creation even if product addition fails
      }
    }

    const result = await readSheet(SHEET_TYPES[req.query.type], dateObj)
    let nextRow = result.length + 4
    nextRow = Math.max(nextRow, 4)

    const sheetName = getMonthlySheetName(SHEET_TYPES[req.query.type], dateObj)

    const range = `${sheetName}!A${nextRow}`

    const values = [
      formatDateForSheet(dateObj),
      customerName,
      productImage ? `=IMAGE("${productImage}")` : '', // Column C - Product Image (keep current structure)
      productName, // Column D - Product Name (keep current structure)
      color, // Column E
      size, // Column F
      quantity, // Column G
      total, // Column H
      status, // Column I
      linkFb, // Column J
      contactInfo, // Column K
      note, // Column L
      productCode || generatedProductCode || '', // Column M - ProductCode
      orderCode || '', // Column N - OrderCode (mã đặt hàng)
      shippingCode || '', // Column O - ShippingCode (mã vận đơn)
    ]

    console.log('Values array:', values)
    console.log('Values length:', values.length)
    console.log('ProductCode value:', productCode)

    // Add customer to KHÁCH HÀNG sheet if not exists
    try {
      if (customerName && customerName.trim()) {
        const customerExists = await checkCustomerExists(customerName.trim())
        if (!customerExists) {
          await addCustomerToSheet(customerName.trim(), contactInfo, linkFb)
          console.log(`Added new customer: ${customerName}`)
        }
      }
    } catch (error) {
      console.error('Error managing customer data:', error)
      // Continue with order creation even if customer addition fails
    }

    const response = await appendSheet(range, values)
    res.json({ message: 'Order added successfully', data: response })
  } catch (error) {
    console.error('Lỗi khi thêm order:', error)
    res.status(500).json({ error: 'Failed to add order 1 to Google Sheet' })
  }
})

// PUT route để update order status
router.put('/status', async (req, res) => {
  try {
    const { rowIndex, status, selectedDate, sheetType } = req.body

    console.log('Update status request:', { rowIndex, status, selectedDate })

    if (rowIndex === undefined || rowIndex === null || !status || !selectedDate) {
      return res
        .status(400)
        .json({ error: 'Missing required fields: rowIndex, status, selectedDate' })
    }

    // Tạo sheet name theo format tháng/năm
    const sheetName = getMonthlySheetName(
      SHEET_TYPES[sheetType],
      new Date(selectedDate.year, selectedDate.month - 1, 1),
    )

    // Cột I (status) ở row index + 4 (vì sheet bắt đầu từ row 4, và rowIndex tính từ 0)
    const range = `${sheetName}!I${rowIndex + 4}`

    console.log('Updating range:', range, 'with status:', status)

    const response = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'RAW',
      requestBody: {
        values: [[status]],
      },
    })

    console.log('Update response:', response.data)

    res.json({
      success: true,
      message: 'Status updated successfully',
      data: response.data,
    })
  } catch (error) {
    console.error('Error updating status:', error)
    res.status(500).json({
      error: 'Failed to update status',
      message: error.message,
    })
  }
})

// PUT route để update toàn bộ order
router.put('/:rowIndex', async (req, res) => {
  try {
    const { rowIndex } = req.params
    const {
      date,
      customerName,
      productCode,
      orderCode,
      shippingCode,
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
      month,
      sheetType,
    } = req.body

    console.log('Update full order request:', { rowIndex, body: req.body })

    if (!rowIndex || rowIndex === undefined) {
      return res.status(400).json({ error: 'Missing rowIndex parameter' })
    }

    // Logic add sản phẩm mới khi update: chỉ khi KHÔNG có productCode và có productImage + productName mới
    const dateObj = date ? parseDateString(date) : new Date()
    let generatedProductCode = '' // Declare variable at function scope

    if (!!(!productCode && productImage)) {
      try {
        // Generate unique product code
        generatedProductCode = generateProductCode() // Use new format: SP{year}{month}{date}{hour}{minute}{second}

        // Add to products sheet
        const existingProducts = await readSheet(SHEET_TYPES.PRODUCTS, dateObj)
        const productNextRow = existingProducts.length + 2
        const productSheetName = getMonthlySheetName(SHEET_TYPES.PRODUCTS, dateObj)
        const productRange = `${productSheetName}!A${productNextRow}`
        const productValues = [
          formatDateForSheet(dateObj),
          generatedProductCode,
          productImage ? `=IMAGE("${productImage}")` : '',
          productName,
        ]
        await appendSheet(productRange, productValues)
        console.log(`Added new product to sheet: ${generatedProductCode}`)
      } catch (error) {
        console.error('Failed to add product to sheet:', error)
        res
          .status(500)
          .json({ error: 'Failed to add product to Google Sheet', errorMSG: error.message })
        return
      }
    }

    // Extract month/year from the order's month field (format: "10/2025")
    let sheetDate = dateObj
    if (month) {
      const [monthStr, yearStr] = month.split('/')
      const selectedDate = {
        month: parseInt(monthStr),
        year: parseInt(yearStr),
      }
      sheetDate = new Date(selectedDate.year, selectedDate.month - 1, 1)
    }

    // Tạo sheet name theo format tháng/năm
    const sheetName = getMonthlySheetName(SHEET_TYPES[sheetType], sheetDate)

    // Row trong sheet (rowIndex + 4 vì sheet bắt đầu từ row 4)
    const targetRow = parseInt(rowIndex) + 4
    const range = `${sheetName}!A${targetRow}:O${targetRow}` // Extend to column N

    console.log('Updating range:', range)

    // Add customer to KHÁCH HÀNG sheet if not exists
    try {
      if (customerName && customerName.trim()) {
        const customerExists = await checkCustomerExists(customerName.trim())
        if (!customerExists) {
          await addCustomerToSheet(customerName.trim(), contactInfo, linkFb)
          console.log(`Added new customer during update: ${customerName}`)
        }
      }
    } catch (error) {
      console.error('Error managing customer data during update:', error)
      // Continue with order update even if customer addition fails
    }

    // Prepare values array matching the sheet structure
    const values = [
      date,
      customerName,
      productImage ? `=IMAGE("${productImage}")` : '', // Column C - Product Image (keep current)
      productName, // Column D - Product Name (keep current)
      color, // Column E
      size, // Column F
      quantity, // Column G
      total, // Column H
      status, // Column I
      linkFb, // Column J
      contactInfo, // Column K
      note, // Column L
      productCode || generatedProductCode || '', // Column M - ProductCode
      orderCode || '', // Column N - OrderCode (mã đặt hàng)
      shippingCode || '', // Column O - ShippingCode (mã vận đơn)
    ]

    const response = await sheets.spreadsheets.values.update({
      spreadsheetId,
      range,
      valueInputOption: 'USER_ENTERED',
      requestBody: {
        values: [values],
      },
    })

    console.log('Update response:', response.data)

    res.json({
      success: true,
      message: 'Order updated successfully',
      data: response.data,
    })
  } catch (error) {
    console.error('Error updating order:', error)
    res.status(500).json({
      error: 'Failed to update order',
      message: error.message,
    })
  }
})

// Debug route to check raw sheet data
router.get('/debug/products', async (req, res) => {
  try {
    const currentDate = new Date()
    const sheetName = getMonthlySheetName(SHEET_TYPES.PRODUCTS, currentDate)

    console.log('Debug - Sheet name:', sheetName)

    const products = await readSheet(SHEET_TYPES.PRODUCTS, currentDate)

    res.json({
      sheetName,
      totalProducts: products.length,
      products: products,
      firstProduct: products.length > 0 ? products[0] : null,
      firstProductKeys: products.length > 0 ? Object.keys(products[0]) : [],
    })
  } catch (error) {
    console.error('Debug error:', error)
    res.status(500).json({
      error: 'Debug failed',
      message: error.message,
    })
  }
})

// Route để lấy tất cả sản phẩm
router.get('/products', async (req, res) => {
  try {
    const currentDate = new Date()
    let allProducts = []

    // Get products from current month
    try {
      const currentMonthProducts = await readSheet(SHEET_TYPES.PRODUCTS, currentDate)
      allProducts = [...currentMonthProducts]
    } catch (error) {
      console.log(
        `No products sheet found for current month: ${currentDate.getMonth() + 1}/${currentDate.getFullYear()}`,
      )
    }

    // Get products from previous months (last 6 months)
    for (let i = 1; i <= 6; i++) {
      const pastDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1)
      try {
        const pastProducts = await readSheet(SHEET_TYPES.PRODUCTS, pastDate)

        // Only add products that don't already exist (avoid duplicates)
        pastProducts.forEach((product) => {
          const exists = allProducts.some(
            (existing) => existing.productCode === product.productCode,
          )
          if (!exists) {
            allProducts.push(product)
          }
        })
      } catch (error) {
        console.log(
          `No products sheet found for ${pastDate.getMonth() + 1}/${pastDate.getFullYear()}`,
        )
      }
    }

    // Sort by product code (newest first) - SP{year}{month}{date}{hour}{minute}{second}
    // Since the new format is chronological, we can compare the timestamp part numerically
    allProducts.sort((a, b) => {
      // Extract the timestamp part from product code (remove 'SP' prefix)
      const timestampA = a.productCode?.replace('SP', '') || '0'
      const timestampB = b.productCode?.replace('SP', '') || '0'

      // Check if both are valid numeric timestamps (14 digits: YYYYMMDDHHMMSS)
      const isValidA = /^\d{14}$/.test(timestampA)
      const isValidB = /^\d{14}$/.test(timestampB)

      // If both are valid timestamps, compare numerically
      if (isValidA && isValidB) {
        return parseInt(timestampB) - parseInt(timestampA) // Newest first
      }

      // If only one is valid, prioritize the valid one
      if (isValidA && !isValidB) return -1 // A comes first
      if (!isValidA && isValidB) return 1 // B comes first

      // If neither is valid timestamp, sort alphabetically
      return b.productCode?.localeCompare(a.productCode || '') || 0
    })

    res.json({
      success: true,
      data: allProducts,
      total: allProducts.length,
    })
  } catch (error) {
    console.error('Error getting all products:', error)
    res.status(500).json({
      success: false,
      error: 'Failed to get products from sheet',
      message: error.message,
    })
  }
})

// Route để tìm kiếm sản phẩm theo mã
router.get('/products/search/:productCode', async (req, res) => {
  try {
    const { productCode } = req.params
    const currentDate = new Date()

    // Tìm trong sheet sản phẩm tháng hiện tại
    const currentMonthProducts = await readSheet(SHEET_TYPES.PRODUCTS, currentDate)

    // Debug: kiểm tra structure của từng product
    currentMonthProducts.forEach((product, index) => {
      console.log(`Product ${index}:`, {
        productCode: product.productCode,
        productName: product.productName,
        keys: Object.keys(product),
      })
    })

    let product = currentMonthProducts.find((p) => p.productCode === productCode)
    console.log('Found product:', product)

    if (!product) {
      // Tìm trong 2 tháng trước
      for (let i = 1; i <= 2; i++) {
        const pastDate = new Date(currentDate.getFullYear(), currentDate.getMonth() - i, 1)
        try {
          const pastProducts = await readSheet(SHEET_TYPES.PRODUCTS, pastDate)
          product = pastProducts.find((p) => p.productCode === productCode)
          if (product) break
        } catch (error) {
          console.log(`No sheet found for ${pastDate.getMonth() + 1}/${pastDate.getFullYear()}`)
        }
      }
    }

    if (product) {
      res.json({
        success: true,
        data: product,
      })
    } else {
      res.json({
        success: false,
        message: 'Không tìm thấy sản phẩm với mã này',
      })
    }
  } catch (error) {
    console.error('Lỗi khi tìm kiếm sản phẩm:', error)
    res.status(500).json({
      success: false,
      error: 'Failed to search product',
    })
  }
})

// Route để thêm sản phẩm mới vào sheet
router.post('/products', async (req, res) => {
  try {
    const { productCode, productImage, productName } = req.body

    if (!productCode || !productImage || !productName) {
      return res.status(400).json({
        error: 'Missing required fields: productCode, productImage, productName',
      })
    }

    const currentDate = new Date()

    // Kiểm tra xem sản phẩm đã tồn tại chưa
    const existingProducts = await readSheet(SHEET_TYPES.PRODUCTS, currentDate)
    const existingProduct = existingProducts.find((p) => p.productCode === productCode)

    if (existingProduct) {
      return res.json({
        success: true,
        message: 'Sản phẩm đã tồn tại',
        data: existingProduct,
      })
    }

    // Thêm sản phẩm mới
    const nextRow = existingProducts.length + 2
    const sheetName = getMonthlySheetName(SHEET_TYPES.PRODUCTS, currentDate)
    const range = `${sheetName}!A${nextRow}`

    const values = [
      formatDateForSheet(currentDate),
      productCode,
      productImage ? `=IMAGE("${productImage}")` : '',
      productName,
    ]

    const response = await appendSheet(range, values)

    res.json({
      success: true,
      message: 'Đã thêm sản phẩm mới vào sheet',
      data: response,
    })
  } catch (error) {
    console.error('Lỗi khi thêm sản phẩm:', error)
    res.status(500).json({
      error: 'Failed to add product to sheet',
    })
  }
})

// Route để lấy danh sách customers
router.get('/customers', async (req, res) => {
  try {
    const customers = await getCustomersFromSheet()

    res.json({
      success: true,
      data: customers,
      total: customers.length,
    })
  } catch (error) {
    console.error('Error getting customers:', error)
    res.status(500).json({
      success: false,
      error: 'Failed to get customers from sheet',
      message: error.message,
    })
  }
})

// Revenue calculation route
router.post('/revenue', async (req, res) => {
  try {
    const { type, year, month } = req.body

    let customerIncome = 0
    let ctvIncome = 0
    let totalExpense = 0
    let details = []

    // Helper function to parse currency string to number
    function parseCurrency(currencyStr) {
      if (!currencyStr || currencyStr === '') return 0
      // Remove all non-digit characters except minus sign
      const cleanStr = currencyStr.toString().replace(/[^\d-]/g, '')
      return parseInt(cleanStr) || 0
    }

    // Helper function to get month name from number
    function getMonthName(monthNum, year) {
      const months = [
        'THÁNG 1',
        'THÁNG 2',
        'THÁNG 3',
        'THÁNG 4',
        'THÁNG 5',
        'THÁNG 6',
        'THÁNG 7',
        'THÁNG 8',
        'THÁNG 9',
        'THÁNG 10',
        'THÁNG 11',
        'THÁNG 12',
      ]
      return `${months[monthNum - 1]} ${year}`
    }

    if (type === 'month') {
      // Get data for specific month

      // 1. Get customer income from "BÁN HÀNG" sheet
      try {
        const customerData = await readSheet(SHEET_TYPES.ORDERS, new Date(year, month - 1, 1))

        customerData.forEach((order) => {
          if (order.total) {
            customerIncome += parseCurrency(order.total)
          }
        })
      } catch (error) {
        console.error('Error fetching customer data:', error)
      }

      // 2. Get CTV income from "CTV" sheet
      try {
        const ctvData = await readSheet(SHEET_TYPES.CTV_ORDERS, new Date(year, month - 1, 1))

        ctvData.forEach((order) => {
          if (order.total) {
            ctvIncome += parseCurrency(order.total)
          }
        })
      } catch (error) {
        console.error('Error fetching CTV data:', error)
      }

      // 3. Get expenses from "ORDCHINA" sheet, cell K2 (total import cost)
      try {
        const chinaSheetName = `ORDCHINA_${month}_${year}`
        const chinaResponse = await sheets.spreadsheets.values.get({
          spreadsheetId,
          range: `${chinaSheetName}!K2`, // Cell K2 contains total import cost
        })

        const chinaRows = chinaResponse.data.values || []
        if (chinaRows.length > 0 && chinaRows[0][0]) {
          totalExpense = parseCurrency(chinaRows[0][0])
        }
      } catch (error) {
        console.error('Error fetching China order data:', error)
      }

      const totalIncome = customerIncome + ctvIncome
      const profit = totalIncome - totalExpense
      const profitMargin = totalIncome > 0 ? Math.round((profit / totalIncome) * 100) : 0

      details = [
        {
          period: `${month}/${year}`,
          customerIncome,
          ctvIncome,
          totalIncome,
          expense: totalExpense,
          profit,
          profitMargin,
        },
      ]
        } else if (type === 'year') {
      // Get data for full year (12 months) - Load all months in parallel
      const months = Array.from({ length: 12 }, (_, i) => i + 1)
      
      // Create all promises for parallel execution
      const monthPromises = months.map(async (m) => {
        let monthCustomerIncome = 0
        let monthCtvIncome = 0
        let monthExpense = 0

        // Load all data for this month in parallel
        const [customerResult, ctvResult, expenseResult] = await Promise.allSettled([
          // Get customer income for this month
          readSheet(SHEET_TYPES.ORDERS, new Date(year, m - 1, 1)).catch(() => []),
          
          // Get CTV income for this month
          readSheet(SHEET_TYPES.CTV_ORDERS, new Date(year, m - 1, 1)).catch(() => []),
          
          // Get expenses for this month from cell K2
          (async () => {
            try {
              const chinaSheetName = `ORDCHINA_${m}_${year}`
              const chinaResponse = await sheets.spreadsheets.values.get({
                spreadsheetId,
                range: `${chinaSheetName}!K2`, // Cell K2 contains total import cost
              })
              const chinaRows = chinaResponse.data.values || []
              return chinaRows.length > 0 && chinaRows[0][0] ? parseCurrency(chinaRows[0][0]) : 0
            } catch (error) {
              return 0
            }
          })()
        ])

        // Process customer income
        if (customerResult.status === 'fulfilled') {
          customerResult.value.forEach((order) => {
            if (order.total) {
              monthCustomerIncome += parseCurrency(order.total)
            }
          })
        }

        // Process CTV income
        if (ctvResult.status === 'fulfilled') {
          ctvResult.value.forEach((order) => {
            if (order.total) {
              monthCtvIncome += parseCurrency(order.total)
            }
          })
        }

        // Process expenses
        if (expenseResult.status === 'fulfilled') {
          monthExpense = expenseResult.value
        }

        const monthTotalIncome = monthCustomerIncome + monthCtvIncome
        const monthProfit = monthTotalIncome - monthExpense
        const monthProfitMargin =
          monthTotalIncome > 0 ? Math.round((monthProfit / monthTotalIncome) * 100) : 0

        return {
          month: m,
          period: `${m}/${year}`,
          customerIncome: monthCustomerIncome,
          ctvIncome: monthCtvIncome,
          totalIncome: monthTotalIncome,
          expense: monthExpense,
          profit: monthProfit,
          profitMargin: monthProfitMargin,
        }
      })

      // Execute all month promises in parallel
      const monthResults = await Promise.all(monthPromises)

      // Sort results by month and calculate totals
      const sortedResults = monthResults.sort((a, b) => a.month - b.month)
      
      let yearlyCustomerIncome = 0
      let yearlyCtvIncome = 0
      let yearlyExpense = 0

      details = sortedResults.map(result => {
        yearlyCustomerIncome += result.customerIncome
        yearlyCtvIncome += result.ctvIncome
        yearlyExpense += result.expense
        
        // Return without the month field
        const { month, ...detailResult } = result
        return detailResult
      })

      customerIncome = yearlyCustomerIncome
      ctvIncome = yearlyCtvIncome
      totalExpense = yearlyExpense
    }

    const totalIncome = customerIncome + ctvIncome
    const totalProfit = totalIncome - totalExpense
    const profitMargin = totalIncome > 0 ? Math.round((totalProfit / totalIncome) * 100) : 0

    const result = {
      totalIncome,
      totalExpense,
      totalProfit,
      profitMargin,
      details,
    }

    res.json(result)
  } catch (error) {
    console.error('Error calculating revenue:', error)
    res.status(500).json({
      error: 'Internal server error',
      message: error.message,
    })
  }
})

function formatDateForSheet(date) {
  // Format dạng: dd/MM/yyyy hoặc ISO nếu bạn config Sheet đọc kiểu khác
  const d = new Date(date)
  return `${d.getDate()}/${d.getMonth() + 1}/${d.getFullYear()}`
}

// Helper function to parse DD/MM/YYYY date format
function parseDateString(dateString) {
  if (!dateString) return new Date()

  // If it's already a valid Date object, return it
  if (dateString instanceof Date) return dateString

  // Check if it's in DD/MM/YYYY format
  const ddmmyyyyPattern = /^(\d{1,2})\/(\d{1,2})\/(\d{4})$/
  const match = dateString.match(ddmmyyyyPattern)

  if (match) {
    const [, day, month, year] = match
    // Create date in local timezone by using Date constructor with separate parameters
    // This avoids timezone issues with ISO string parsing
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day), 12, 0, 0, 0)
  }

  // Check if it's in YYYY/MM/DD format
  const yyyymmddPattern = /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/
  const match2 = dateString.match(yyyymmddPattern)

  if (match2) {
    const [, year, month, day] = match2
    // Create date in local timezone by using Date constructor with separate parameters
    return new Date(parseInt(year), parseInt(month) - 1, parseInt(day), 12, 0, 0, 0)
  }

  // Try default Date parsing (for other formats)
  const parsedDate = new Date(dateString)

  // If parsing failed, return current date
  if (isNaN(parsedDate.getTime())) {
    console.warn(`Invalid date format: ${dateString}, using current date`)
    return new Date()
  }

  return parsedDate
}

// Đọc dữ liệu từ sheet
async function readSheet(baseSheetName, date) {
  try {
    const sheetName = getMonthlySheetName(baseSheetName, date)

    const spreadsheetId = process.env.GOOGLE_SHEET_ID

    if (!spreadsheetId) {
      throw new Error('GOOGLE_SHEET_ID is not defined')
    }

    const start = new Date()

    const response = await sheets.spreadsheets.get({
      spreadsheetId,
      ranges: [`${sheetName}!A:Z`],
      includeGridData: true,
      fields: 'sheets.data.rowData.values.userEnteredValue',
    })
    const end = new Date()
    const durationMs = end.getTime() - start.getTime() // thời gian chạy (ms)
    console.log(`Time taken get data gg: ${durationMs} ms`)

    const rows = response.data.sheets?.[0]?.data?.[0]?.rowData || []

    // Map the data based on the sheet structure
    if (baseSheetName === SHEET_TYPES.ORDERS || baseSheetName === SHEET_TYPES.CTV_ORDERS) {
      // Skip the first 3 rows (headers)
      return rows
        .slice(3)
        .map((row, index) => {
          const cells = row.values || []
          return {
            rowIndex: index,
            date: parseGoogleSheetDate(cells[0]),
            customerName: getCellString(cells[1]),
            productImage: extractImageUrl(getCellString(cells[2])), // Column C - Product Image (current)
            productName: getCellString(cells[3]), // Column D - Product Name (current)
            color: getCellString(cells[4]), // Column E
            size: getCellString(cells[5]), // Column F
            quantity: getCellString(cells[6]), // Column G
            total: getCellString(cells[7]), // Column H
            status: getCellString(cells[8]), // Column I
            linkFb: getCellString(cells[9]), // Column J
            contactInfo: getCellString(cells[10]), // Column K
            note: getCellString(cells[11]), // Column L
            productCode: getCellString(cells[12]) || '', // Column M - Product Code
            orderCode: getCellString(cells[13]) || '', // Column N - Order Code (mã đặt hàng)
            shippingCode: getCellString(cells[14]) || '', // Column O - Shipping Code (mã vận đơn)
            month: `${date.getMonth() + 1}/${date.getFullYear()}`,
          }
        })
        .filter((item) => item.customerName) // Filter out empty rows
    } else if (baseSheetName === SHEET_TYPES.INVENTORY) {
      // Skip the first row (header)
      return rows
        .slice(1)
        .map((row, index) => {
          const cells = row.values || []
          const rowData = {
            rowIndex: index,
            date: parseGoogleSheetDate(cells[0]),
            customerName: getCellString(cells[1]),
            productImage: extractImageUrl(getCellString(cells[2])), // Column C - Product Image (current)
            productName: getCellString(cells[3]), // Column D - Product Name (current)
            color: getCellString(cells[4]), // Column E
            size: getCellString(cells[5]), // Column F
            quantity: getCellString(cells[6]), // Column G
            total: getCellString(cells[7]), // Column H
            status: getCellString(cells[8]), // Column I
            linkFb: getCellString(cells[9]), // Column J
            contactInfo: getCellString(cells[10]), // Column K
            note: getCellString(cells[11]), // Column L
            productCode: getCellString(cells[12]) || '', // Column M - Product Code
            orderCode: getCellString(cells[13]) || '', // Column N - Order Code (mã đặt hàng)
            shippingCode: getCellString(cells[14]) || '', // Column O - Shipping Code (mã vận đơn)
            month: `${date.getMonth() + 1}/${date.getFullYear()}`,
          }

          // Debug log for shipping code - only for ORDERS
          if (baseSheetName === SHEET_TYPES.ORDERS && getCellString(cells[14])) {
            console.log(
              `[ORDERS] Row ${index + 2}: Found shipping code "${getCellString(cells[14])}" for customer "${getCellString(cells[1])}"`,
            )
          }

          return rowData
        })
        .filter((item) => item.productName) // Filter out empty rows
    } else if (baseSheetName === SHEET_TYPES.PRODUCTS) {
      // Skip the first row (header)
      return rows
        .slice(1)
        .map((row, index) => {
          const cells = row.values || []
          return {
            rowIndex: index,
            date: parseGoogleSheetDate(cells[0]),
            productCode: getCellString(cells[1]),
            productImage: extractImageUrl(getCellString(cells[2])),
            productName: getCellString(cells[3]),
            month: `${date.getMonth() + 1}/${date.getFullYear()}`,
          }
        })
        .filter((item) => item.productCode) // Filter out empty rows
    }
    return []
  } catch (error) {
    console.error('Lỗi khi đọc dữ liệu:', error)
    throw error
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
    })
    return response.data
  } catch (error) {
    console.error('Lỗi khi ghi dữ liệu:', error)
    throw error
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
    })
    return response.data
  } catch (error) {
    console.error('Lỗi khi thêm dữ liệu:', error)
    throw error
  }
}

function extractImageUrl(cellValue) {
  const match = cellValue.match(/^=IMAGE\("([^"]+)"\)/)

  return match ? match[1] : ''
}

function getCellString(cell) {
  const val = cell?.userEnteredValue
  if (!val) return ''

  if (val.stringValue !== undefined) return val.stringValue
  if (val.numberValue !== undefined) return String(val.numberValue)
  if (val.boolValue !== undefined) return String(val.boolValue)
  if (val.formulaValue !== undefined) return val.formulaValue

  return ''
}

function parseGoogleSheetDate(cell) {
  const value = cell?.userEnteredValue
  if (!value) return ''
  if (value.numberValue !== undefined) {
    // Google Sheets date serial number -> JS Date
    const baseDate = new Date(Date.UTC(1899, 11, 30)) // 1899-12-30
    const date = new Date(baseDate.getTime() + value.numberValue * 24 * 60 * 60 * 1000)
    return `${date.getDate()}/${date.getMonth() + 1}`
  } else if (value.stringValue !== undefined) {
    return value.stringValue
  } else {
    return ''
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
  ORDERS: 'BÁN HÀNG',
  INVENTORY: 'NHẬP HÀNG',
  CTV_ORDERS: 'CTV',
  PRODUCTS: 'SP',
  CUSTOMERS: 'KHÁCH HÀNG',
}

// Create ORDCHINA record
router.post('/ordchina', async (req, res) => {
  try {
    const {
      managementCode,
      productName,
      productImage,
      status,
      shippingCodes,
      note,
      orderDate,
      quantity,
      importPrice,
      date,
    } = req.body

    const sheetName = `ORDCHINA_${date.month}_${date.year}`

    // Create sheet if it doesn't exist
    await createSheetIfNotExists(sheetName)

    // Get existing data to find next row
    const existingData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:J`,
    })

    const rows = existingData.data.values || []
    const nextRow = rows.length + 1

    // Prepare row data - start from column A
    const rowData = [
      managementCode, // Column A: Mã quản lý order
      productName, // Column B: Tên sản phẩm
      productImage ? `=IMAGE("${productImage}")` : '', // Column C: HÌNH ẢNH
      status, // Column D: STATUS
      shippingCodes, // Column E: MÃ VẬN ĐƠN
      note, // Column F: NOTE
      orderDate, // Column G: NGÀY CHỐT MUA
      '', // Column H: NGÀY Hàng về (empty)
      quantity, // Column I: Số lượng
      importPrice, // Column J: Giá nhập
    ]

    // Insert data at specific row
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${nextRow}:J${nextRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [rowData],
      },
    })

    res.json({ success: true, managementCode })
  } catch (error) {
    console.error('Error creating ORDCHINA record:', error)
    res.status(500).json({ error: error.message })
  }
})

// Helper function to create sheet if it doesn't exist
async function createSheetIfNotExists(sheetName) {
  try {
    // Check if sheet exists
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId,
    })

    const sheetExists = spreadsheet.data.sheets.some(
      (sheet) => sheet.properties.title === sheetName,
    )

    if (!sheetExists) {
      // Create new sheet
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

      // Add headers - start from column A
      const headers = [
        'Mã quản lý order', // Column A
        'Tên sản phẩm', // Column B
        'HÌNH ẢNH', // Column C
        'STATUS', // Column D
        'MÃ VẬN ĐƠN', // Column E
        'NOTE', // Column F
        'NGÀY CHỐT MUA', // Column G
        'NGÀY Hàng về', // Column H
        'Số lượng', // Column I
        'Giá nhập', // Column J
      ]

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1:J1`,
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

// Customer management functions
async function checkCustomerExists(customerName) {
  try {
    const sheetName = SHEET_TYPES.CUSTOMERS

    // Get existing customers data
    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:C`,
    })

    const rows = response.data.values || []

    // Skip header row and check if customer exists
    return rows
      .slice(1)
      .some((row) => row[0] && row[0].toLowerCase().trim() === customerName.toLowerCase().trim())
  } catch (error) {
    console.error('Error checking customer existence:', error)
    // If sheet doesn't exist, customer doesn't exist
    return false
  }
}

async function addCustomerToSheet(customerName, contactInfo, linkFb) {
  try {
    const sheetName = SHEET_TYPES.CUSTOMERS

    // Create sheet if it doesn't exist
    await createCustomerSheetIfNotExists()

    // Get existing data to find next row
    const existingData = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:C`,
    })

    const rows = existingData.data.values || []
    const nextRow = rows.length + 1

    // Prepare customer data
    const customerData = [
      customerName || '', // Column A: Tên khách hàng
      contactInfo || '', // Column B: Địa chỉ/SDT
      linkFb || '', // Column C: Link FB
    ]

    // Insert customer data
    await sheets.spreadsheets.values.update({
      spreadsheetId,
      range: `${sheetName}!A${nextRow}:C${nextRow}`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: [customerData],
      },
    })

    console.log(`Added new customer to sheet: ${customerName}`)
    return true
  } catch (error) {
    console.error('Error adding customer to sheet:', error)
    throw error
  }
}

async function createCustomerSheetIfNotExists() {
  try {
    const sheetName = SHEET_TYPES.CUSTOMERS

    // Check if sheet exists
    const spreadsheet = await sheets.spreadsheets.get({
      spreadsheetId,
    })

    const sheetExists = spreadsheet.data.sheets.some(
      (sheet) => sheet.properties.title === sheetName,
    )

    if (!sheetExists) {
      // Create new sheet
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
      const headers = [
        'Tên khách hàng', // Column A
        'Địa chỉ/SDT', // Column B
        'Link FB', // Column C
      ]

      await sheets.spreadsheets.values.update({
        spreadsheetId,
        range: `${sheetName}!A1:C1`,
        valueInputOption: 'USER_ENTERED',
        resource: {
          values: [headers],
        },
      })
    }
  } catch (error) {
    console.error('Error creating customer sheet:', error)
    throw error
  }
}

async function getCustomersFromSheet() {
  try {
    const sheetName = SHEET_TYPES.CUSTOMERS

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:C`,
    })

    const rows = response.data.values || []

    // Skip header row and map to customer objects
    return rows
      .slice(1)
      .map((row, index) => ({
        rowIndex: index,
        customerName: row[0] || '',
        contactInfo: row[1] || '',
        linkFb: row[2] || '',
      }))
      .filter((customer) => customer.customerName) // Filter out empty rows
  } catch (error) {
    console.error('Error getting customers from sheet:', error)
    return []
  }
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
//       if (baseSheetName === SHEET_TYPES.ORDERS || baseSheetName === SHEET_TYPES.CTV_ORDERS) {
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