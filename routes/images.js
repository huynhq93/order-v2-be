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
const upload = multer({
  storage,
  limits: {
    fileSize: 10 * 1024 * 1024, // 10MB limit
  },
  fileFilter: (req, file, cb) => {
    // Only allow image files
    if (file.mimetype.startsWith('image/')) {
      cb(null, true)
    } else {
      cb(new Error('Only image files are allowed!'), false)
    }
  },
})

router.post('/', upload.single('file'), async (req, res) => {
  try {
    const file = req.file
    console.log(
      'Received file:',
      file
        ? {
            originalname: file.originalname,
            mimetype: file.mimetype,
            size: file.size,
          }
        : 'No file received',
    )

    if (!file) {
      return res.status(400).json({
        error: 'No image file provided',
        message: 'Please select an image file to upload',
      })
    }

    // Convert buffer to base64
    const base64 = file.buffer.toString('base64')
    const dataUri = `data:${file.mimetype};base64,${base64}`

    const result = await cloudinary.uploader.upload(dataUri, {
      folder: 'orders',
    })

    console.log('Upload successful:', result.secure_url)
    res.json({ url: result.secure_url })
  } catch (err) {
    console.error('Upload error:', err)
    res.status(500).json({
      error: 'Upload failed',
      message: err.message || 'An error occurred while uploading the image',
    })
  }
})

module.exports = router