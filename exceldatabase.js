const express = require('express');
const multer = require('multer');
const xlsx = require('xlsx');
const mongoose = require('mongoose');
const Schema = mongoose.Schema;
const app = express();
const port = 3000;
var fs = require('fs');

// Connect to MongoDB using "127.0.0.1:27017" (replace with your database name)
mongoose.connect('mongodb://127.0.0.1:27017/excel_upload_final', {
  useNewUrlParser: true,
  useUnifiedTopology: true
});

// Define a Mongoose schema for dynamic Excel data
const dynamicExcelSchema = new Schema({});

// Create a Mongoose model based on the dynamic schema
const DynamicExcelModel = mongoose.model('DynamicExcel', dynamicExcelSchema);

// Function to create a dynamic schema based on Excel headings
function createDynamicSchema(headings) {
  const dynamicSchema = {};

  headings.forEach((heading) => {
    dynamicSchema[heading] = String; // You can assume all fields are strings, but adjust data types as needed
  });

  return dynamicSchema;
}

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

  // Extract the headings (first row in the Excel sheet)
  const headings = Object.keys(sheetData[0]);

  // Create a dynamic schema based on headings for Excel data
  const dynamicSchema = createDynamicSchema(headings);

  // Create a new Mongoose model using the dynamic schema
  const ExcelDataModel = mongoose.model('ExcelData', new Schema(dynamicSchema));

  // Create a new MongoDB document with the form data and Excel data
  const formData = {
    legalEntity: req.body.legalEntity,
    website: req.body.website,
    authorizedPerson: req.body.authorizedPerson,
    email: req.body.email,
    excelData: sheetData // Store the Excel data as is
  };

  try {
    // Insert the form data into MongoDB
    const document = new ExcelDataModel(formData);
    await document.save();
    
    let textdata=JSON.stringify(formData)
fs.writeFile('mynewfile3.json', textdata, function (err) {
  if (err) throw err;
  console.log('Replaced!');
});
    // console.log(okdata);
    res.redirect('/view');
  } catch (err) {
    console.error(err);
  
    // Log validation errors
    if (err.errors) {
      console.log('Validation Errors:', err.errors);
    }
  
    res.send('Error uploading file');
  }
  
});

// Route to render the view page
app.get('/view', (req, res) => {
  // Read the JSON file
  fs.readFile('mynewfile3.json', 'utf8', (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Error reading JSON file');
    } else {
      // Parse the JSON data
      const jsonData = JSON.parse(data);
      res.render('view', { jsonData }); // Render the view template and pass the JSON data
    }
  });
});


app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
