const express = require('express')
const cors = require('cors')
require('dotenv').config()

const sheetRoutes = require('./routes/sheets')
const imageRoutes = require('./routes/images')

const app = express()
app.use(cors())
app.use(express.json())

app.use('/api/sheets', sheetRoutes)
app.use('/api/images', imageRoutes)

// Also mount routes without /api prefix for direct access
app.use('/sheets', sheetRoutes)
app.use('/images', imageRoutes)

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
      sheets: '/sheets?type=ORDERS&month=5&year=2025',
      images: '/images/upload',
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