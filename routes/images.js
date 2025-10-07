const express = require('express')
const multer = require('multer')
const { v2: cloudinary } = require('cloudinary')

const router = express.Router()

cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key:    process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
})


const storage = multer.memoryStorage()
const upload = multer({ storage })

router.post('/upload', upload.single('file'), async (req, res) => {
  try {
    const file = req.file
    // console.log('file:',file)
    // const { file } = req.body
    // console.log('file1:', file)

    if (!file) {
      return res.status(400).send('Image is null')
    }

    // Convert buffer to base64
    const base64 = file.buffer.toString('base64')
    const dataUri = `data:${file.mimetype};base64,${base64}`

    const result = await cloudinary.uploader.upload(dataUri, {
      folder: 'orders',
    })

    res.json({ url: result.secure_url })
  } catch (err) {
    console.error(err)
    res.status(500).send('Upload image error')
  }
})

module.exports = router