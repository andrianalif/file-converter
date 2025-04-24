# File Converter & WordPress Publisher

A modern web application that converts Excel, Word, and PDF files to HTML and publishes them to WordPress.

## Features

- Drag and drop file upload
- Support for Excel (.xlsx), Word (.docx), and PDF files
- Real-time preview of converted content
- Direct publishing to WordPress
- Modern, responsive UI
- Fast conversion and publishing

## Prerequisites

- Node.js (v14 or higher)
- Python (v3.8 or higher)
- WordPress site with REST API enabled
- WordPress Application Password

## Setup

1. Install frontend dependencies:
```bash
npm install
```

2. Install backend dependencies:
```bash
pip install -r requirements.txt
```

3. Create a `.env` file in the root directory with your WordPress credentials:
```
WP_USER=your_wordpress_username
WP_APP_PASSWORD=your_application_password
WP_API_URL=your_wordpress_api_url
```

## Running the Application

1. Start the backend server:
```bash
python server.py
```

2. Start the frontend development server:
```bash
npm start
```

3. Open your browser and navigate to `http://localhost:3000`

## Usage

1. Drag and drop a file (Excel, Word, or PDF) into the upload area
2. The file name will be automatically set as the page title (you can modify it)
3. Click "Convert" to preview the converted content
4. Click "Publish to WordPress" to publish the content to your WordPress site
5. The published page URL will be displayed in a success message

## Security Notes

- Never commit your `.env` file to version control
- Keep your WordPress Application Password secure
- The application uses CORS to protect the backend API

## Technologies Used

- Frontend: React, TypeScript, Material-UI
- Backend: Python Flask
- File Processing: openpyxl, python-docx, PyPDF2
- WordPress Integration: WordPress REST API 