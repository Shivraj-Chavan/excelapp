const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const app = express();
const port = 3000;

// Set up Multer for file uploads
const storage = multer.memoryStorage();
const upload = multer({ storage: storage });

// Set up EJS as your view engine (replace with your preferred view engine)
app.set('view engine', 'ejs');

// Define routes

// Route to the homepage with the file upload form
app.get('/', (req, res) => {
  res.render('index'); // Replace with the name of your HTML form page
});

// Route to handle file upload
app.post('/upload', upload.single('file'), async (req, res) => {
  const fileBuffer = req.file.buffer;
  const workbook = xlsx.read(fileBuffer, { type: 'buffer' });
  const sheetName = workbook.SheetNames[0];
  const sheetData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetName]);

  
  // Create a new MongoDB document with the form data and Excel data
  const formData = {
    legalEntity: req.body.legalEntity,
    website: req.body.website,
    authorizedPerson: req.body.authorizedPerson,
    email: req.body.email,
    excelData: sheetData 
  };
  res.render('view', { formData });
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
