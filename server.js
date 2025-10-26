const express = require('express')
const cors = require('cors')
require('dotenv').config()

const sheetRoutes = require('./routes/sheets')
const imageRoutes = require('./routes/images')
const { router: authRoutes, verifyToken } = require('./routes/auth')

const app = express()
app.use(cors())

// Increase payload limits for image uploads
app.use(express.json({ limit: '50mb' }))
app.use(express.urlencoded({ limit: '50mb', extended: true }))

// Auth routes (no token required)
app.use('/api/auth', authRoutes)

// Protected routes (require token)
app.use('/api/sheets', verifyToken, sheetRoutes)
app.use('/api/images', verifyToken, imageRoutes)

// Also mount routes without /api prefix for direct access (protected)
app.use('/sheets', verifyToken, sheetRoutes)
app.use('/images', verifyToken, imageRoutes)

app.get('/api/test', (req, res) => {
  res.json({
    message: 'Dữ liệu nhận được',
  })
})

app.get('/test', (req, res) => {
  res.json({
    message: 'Backend is working!',
  })
})

// Root endpoint
app.get('/', (req, res) => {
  res.json({
    message: 'Order Management Backend API',
    endpoints: {
      auth: {
        login: '/api/auth/login (POST)',
        verify: '/api/auth/verify (POST)',
        init: '/api/auth/init-accounts (POST)'
      },
      sheets: '/sheets?type=ORDERS&month=5&year=2025 (Protected)',
      images: '/images/upload (Protected)',
      revenue: '/sheets/revenue (POST, Protected)',
      test: '/test',
    },
  })
})

// For Vercel serverless
module.exports = app

// For local development
if (process.env.NODE_ENV !== 'production') {
  app.listen(5176, () => console.log('✅ Backend running on http://localhost:5176'))
}