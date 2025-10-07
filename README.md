# Order Management System - Backend

Backend API server for Google Sheets order management system.

## Features

- Google Sheets integration
- Image upload with Cloudinary
- Order status management
- CORS enabled for frontend integration

## Environment Variables

```
GOOGLE_CLIENT_EMAIL=your_service_account_email
GOOGLE_PRIVATE_KEY=your_private_key
GOOGLE_SHEET_ID=your_sheet_id
CLOUDINARY_CLOUD_NAME=your_cloud_name
CLOUDINARY_API_KEY=your_api_key
CLOUDINARY_API_SECRET=your_api_secret
```

## Deployment

Deploy to Vercel:

1. Push to GitHub
2. Connect to Vercel
3. Set environment variables
4. Deploy

## Local Development

```bash
npm install
npm start
```

Server runs on http://localhost:5176
