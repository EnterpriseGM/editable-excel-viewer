
const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');
const path = require('path');
const fs = require('fs');

const app = express();
const PORT = process.env.PORT || 3001;

// Middleware
app.use(cors());
app.use(express.json());

// Configure multer for file uploads
const storage = multer.diskStorage({
  destination: (req, file, cb) => {
    const uploadDir = path.join(__dirname, 'uploads');
    if (!fs.existsSync(uploadDir)) {
      fs.mkdirSync(uploadDir, { recursive: true });
    }
    cb(null, uploadDir);
  },
  filename: (req, file, cb) => {
    cb(null, `${Date.now()}-${file.originalname}`);
  }
});

const upload = multer({ 
  storage,
  fileFilter: (req, file, cb) => {
    const filetypes = /xlsx|xls/;
    const extname = filetypes.test(path.extname(file.originalname).toLowerCase());
    const mimetype = filetypes.test(file.mimetype);
    
    if (mimetype && extname) {
      return cb(null, true);
    } else {
      cb(new Error('Only Excel files are allowed!'));
    }
  }
});

// Store uploaded files info
const uploadedFiles = {};

// API endpoints
// Upload Excel file
app.post('/api/upload', upload.single('file'), (req, res) => {
  try {
    if (!req.file) {
      return res.status(400).json({ error: 'No file uploaded' });
    }

    const fileId = req.file.filename;
    const filePath = req.file.path;
    
    // Read the Excel file
    const workbook = XLSX.readFile(filePath);
    const sheetNames = workbook.SheetNames;
    
    if (sheetNames.length === 0) {
      return res.status(400).json({ error: 'Invalid Excel file - no sheets found' });
    }
    
    const activeSheet = sheetNames[0];
    const sheets = {};
    
    // Process all sheets
    for (const sheetName of sheetNames) {
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });
      
      if (jsonData.length > 0) {
        // Extract headers (first row) and the rest of the data
        const headers = jsonData[0] || [];
        const rows = jsonData.slice(1) || [];
        
        sheets[sheetName] = {
          headers,
          data: rows
        };
      } else {
        // Handle empty sheets
        sheets[sheetName] = {
          headers: [],
          data: []
        };
      }
    }
    
    // Store the parsed data
    uploadedFiles[fileId] = {
      originalName: req.file.originalname,
      filePath,
      headers: sheets[activeSheet]?.headers || [],
      data: sheets[activeSheet]?.data || [],
      sheetNames,
      activeSheet,
      sheets
    };
    
    // Return the data to the client
    return res.status(200).json({
      fileId,
      fileName: req.file.originalname,
      headers: sheets[activeSheet]?.headers || [],
      data: sheets[activeSheet]?.data || [],
      sheetNames,
      activeSheet,
      sheets
    });
  } catch (error) {
    console.error('Error processing Excel file:', error);
    return res.status(500).json({ error: 'Failed to process the Excel file' });
  }
});

// Get sheet data
app.get('/api/sheets/:fileId', (req, res) => {
  const { fileId } = req.params;
  
  if (!uploadedFiles[fileId]) {
    return res.status(404).json({ error: 'File not found' });
  }
  
  return res.status(200).json(uploadedFiles[fileId]);
});

// Change active sheet
app.post('/api/sheets/:fileId/change-sheet', (req, res) => {
  const { fileId } = req.params;
  const { sheetName } = req.body;
  
  if (!uploadedFiles[fileId]) {
    return res.status(404).json({ error: 'File not found' });
  }
  
  const fileData = uploadedFiles[fileId];
  
  if (!fileData.sheetNames.includes(sheetName)) {
    return res.status(400).json({ error: 'Sheet not found' });
  }
  
  // Save current sheet's data
  const currentSheetData = {
    headers: fileData.headers,
    data: fileData.data
  };
  
  const updatedSheets = {
    ...fileData.sheets,
    [fileData.activeSheet]: currentSheetData
  };
  
  // Load selected sheet's data
  const newSheetData = updatedSheets[sheetName];
  
  // Update the file data
  uploadedFiles[fileId] = {
    ...fileData,
    headers: newSheetData.headers,
    data: newSheetData.data,
    activeSheet: sheetName,
    sheets: updatedSheets
  };
  
  return res.status(200).json(uploadedFiles[fileId]);
});

// Update cell value
app.post('/api/sheets/:fileId/update-cell', (req, res) => {
  const { fileId } = req.params;
  const { rowIndex, columnIndex, value } = req.body;
  
  if (!uploadedFiles[fileId]) {
    return res.status(404).json({ error: 'File not found' });
  }
  
  const fileData = uploadedFiles[fileId];
  const newData = [...fileData.data];
  
  // If the row doesn't exist, create it
  if (!newData[rowIndex]) {
    newData[rowIndex] = Array(fileData.headers.length).fill('');
  }
  
  // Update the value
  newData[rowIndex][columnIndex] = value;
  
  // Update the sheets object with the modified data
  const updatedSheets = {
    ...fileData.sheets,
    [fileData.activeSheet]: {
      headers: fileData.headers,
      data: newData
    }
  };
  
  // Update the file data
  uploadedFiles[fileId] = {
    ...fileData,
    data: newData,
    sheets: updatedSheets
  };
  
  return res.status(200).json(uploadedFiles[fileId]);
});

// Export Excel file
app.get('/api/sheets/:fileId/export', (req, res) => {
  const { fileId } = req.params;
  
  if (!uploadedFiles[fileId]) {
    return res.status(404).json({ error: 'File not found' });
  }
  
  const fileData = uploadedFiles[fileId];
  
  try {
    // Create a new workbook
    const workbook = XLSX.utils.book_new();
    
    // Add each sheet to the workbook
    for (const sheetName of fileData.sheetNames) {
      const sheetData = sheetName === fileData.activeSheet 
        ? { headers: fileData.headers, data: fileData.data } 
        : fileData.sheets[sheetName];
      
      if (sheetData) {
        // Create a worksheet from the data
        const worksheet = XLSX.utils.aoa_to_sheet([sheetData.headers, ...sheetData.data]);
        
        // Add the worksheet to the workbook
        XLSX.utils.book_append_sheet(workbook, worksheet, sheetName);
      }
    }
    
    // Create a temporary file for the export
    const exportDir = path.join(__dirname, 'exports');
    if (!fs.existsSync(exportDir)) {
      fs.mkdirSync(exportDir, { recursive: true });
    }
    
    const exportFileName = `export-${Date.now()}.xlsx`;
    const exportPath = path.join(exportDir, exportFileName);
    
    // Write the workbook to a file
    XLSX.writeFile(workbook, exportPath);
    
    // Send the file to the client
    res.download(exportPath, fileData.originalName, (err) => {
      if (err) {
        console.error('Error sending file:', err);
      }
      
      // Delete the temporary file after sending
      fs.unlink(exportPath, (unlinkErr) => {
        if (unlinkErr) {
          console.error('Error deleting temporary file:', unlinkErr);
        }
      });
    });
  } catch (error) {
    console.error('Error exporting Excel file:', error);
    return res.status(500).json({ error: 'Failed to export the Excel file' });
  }
});

// Start the server
app.listen(PORT, () => {
  console.log(`Server running on port ${PORT}`);
});
