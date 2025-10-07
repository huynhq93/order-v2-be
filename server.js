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
app.get('/api/test', (req, res) => {
    res.json({
      message: 'Dữ liệu nhận được',
    });
  });

app.listen(5176, () => console.log('✅ Backend running on http://localhost:5176'))