const express = require('express')
const { google } = require('googleapis')
const jwt = require('jsonwebtoken')

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
const JWT_SECRET = process.env.JWT_SECRET || 'your-secret-key-change-this-in-production'
const JWT_EXPIRES_IN = process.env.JWT_EXPIRES_IN || '30d'

// Login endpoint
router.post('/login', async (req, res) => {
  try {
    const { username, password } = req.body

    if (!username || !password) {
      return res.status(400).json({
        success: false,
        message: 'Username và password là bắt buộc',
      })
    }

    // Get accounts from Google Sheet
    const accounts = await getAccountsFromSheet()

    // Find user
    const user = accounts.find((account) => account.username === username)

    if (!user) {
      return res.status(401).json({
        success: false,
        message: 'Tài khoản không tồn tại',
      })
    }

    // Check password (simple comparison for now, can be enhanced with bcrypt)
    if (user.password !== password) {
      return res.status(401).json({
        success: false,
        message: 'Mật khẩu không đúng',
      })
    }

    // Generate JWT token
    const token = jwt.sign(
      {
        userId: user.username,
        role: user.role,
      },
      JWT_SECRET,
      { expiresIn: JWT_EXPIRES_IN },
    )

    res.json({
      success: true,
      message: 'Đăng nhập thành công',
      data: {
        token,
        user: {
          username: user.username,
          role: user.role,
        },
      },
    })
  } catch (error) {
    console.error('Login error:', error)
    res.status(500).json({
      success: false,
      message: 'Lỗi server',
    })
  }
})

// Verify token endpoint
router.post('/verify', async (req, res) => {
  try {
    const { token } = req.body

    if (!token) {
      return res.status(400).json({
        success: false,
        message: 'Token là bắt buộc',
      })
    }

    const decoded = jwt.verify(token, JWT_SECRET)

    res.json({
      success: true,
      data: {
        user: {
          username: decoded.userId,
          role: decoded.role,
        },
      },
    })
  } catch (error) {
    res.status(401).json({
      success: false,
      message: 'Token không hợp lệ',
    })
  }
})

// Get accounts from Google Sheet
async function getAccountsFromSheet() {
  try {
    const sheetName = 'Account'

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId,
      range: `${sheetName}!A:C`,
    })

    const rows = response.data.values || []

    // Skip header row and map to account objects
    return rows
      .slice(1)
      .map((row) => ({
        username: row[0] || '',
        password: row[1] || '',
        role: row[2] || 'nv',
      }))
      .filter((account) => account.username) // Filter out empty rows
  } catch (error) {
    console.error('Error getting accounts from sheet:', error)
    return []
  }
}

// Middleware to verify JWT token
function verifyToken(req, res, next) {
  const authHeader = req.headers['authorization']
  const token = authHeader && authHeader.split(' ')[1] // Bearer TOKEN

  if (!token) {
    return res.status(401).json({
      success: false,
      message: 'Access token required',
    })
  }

  jwt.verify(token, JWT_SECRET, (err, user) => {
    if (err) {
      return res.status(403).json({
        success: false,
        message: 'Invalid token',
      })
    }
    req.user = user
    next()
  })
}

// Initialize accounts (add default accounts)
router.post('/init-accounts', async (req, res) => {
  try {
    const sheetName = 'Account'

    // Check if sheet exists and has data
    const existingAccounts = await getAccountsFromSheet()

    if (existingAccounts.length > 0) {
      return res.json({
        success: true,
        message: 'Accounts already exist',
        data: existingAccounts.map((acc) => ({ username: acc.username, role: acc.role })),
      })
    }

    // Create sheet if it doesn't exist and add default accounts
    await createAccountSheetIfNotExists()

    // Add default accounts
    const defaultAccounts = [
      ['admin', 'admin2808', 'admin'],
      ['nv001', 'nv001', 'nv'],
    ]

    await sheets.spreadsheets.values.append({
      spreadsheetId,
      range: `${sheetName}!A:C`,
      valueInputOption: 'USER_ENTERED',
      resource: {
        values: defaultAccounts,
      },
    })

    res.json({
      success: true,
      message: 'Default accounts created successfully',
      data: defaultAccounts.map((acc) => ({ username: acc[0], role: acc[2] })),
    })
  } catch (error) {
    console.error('Error initializing accounts:', error)
    res.status(500).json({
      success: false,
      message: 'Failed to initialize accounts',
    })
  }
})

// Helper function to create Account sheet if it doesn't exist
async function createAccountSheetIfNotExists() {
  try {
    const sheetName = 'Account'

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
      const headers = ['Username', 'Password', 'Role']

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
    console.error('Error creating Account sheet:', error)
    throw error
  }
}

module.exports = { router, verifyToken }
