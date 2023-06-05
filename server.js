const express = require('express');
const bodyParser = require('body-parser');
const multer = require('multer');
const ExcelJS = require('exceljs');

const app = express();
const port = process.env.PORT || 3000;
const path = require('path');
const fs = require('fs');

app.use(bodyParser.urlencoded({ extended: false }));
app.use(bodyParser.json());

app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'index.html'));
});

const storage = multer.diskStorage({
  filename(req, file, cb) {
    cb(null, file.originalname);
  },
});

const upload = multer({ storage });

const workbook = new ExcelJS.Workbook();
let worksheet;

function loadStudentsFromExcel(file) {
  const newWorkbook = new ExcelJS.Workbook();
  newWorkbook.xlsx.readFile(file)
    .then(() => {
      const newWorksheet = newWorkbook.getWorksheet();

      newWorksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
          worksheet.getCell(rowNumber, colNumber).value = cell.value;
        });
      });
      console.log('Students loaded from Excel file successfully!');
      return workbook.xlsx.writeFile('students.xlsx');
    })
    .catch((error) => {
      console.error('Error:', error);
    });
}

function createWorksheet() {
  worksheet = workbook.addWorksheet('Students');

  worksheet.columns = [
    { header: 'Name', key: 'name', width: 20 },
    { header: 'Exam 1', key: 'exam1', width: 10 },
    { header: 'Exam 2', key: 'exam2', width: 10 },
    { header: 'Exam 3', key: 'exam3', width: 10 },
  ];
}

function saveStudentsToExcel() {
  workbook.xlsx.writeFile('students.xlsx')
    .then(() => {
      console.log('Students saved to Excel file successfully!');
    })
    .catch((error) => {
      console.error('Error:', error);
    });
}

app.post('/register', (req, res) => {
  const {
    name, exam1, exam2, exam3,
  } = req.body;

  const student = {
    name,
    exam1: parseFloat(exam1),
    exam2: parseFloat(exam2),
    exam3: parseFloat(exam3),
  };

  worksheet.addRow(student).commit();

  saveStudentsToExcel();

  res.send('Student added successfully!');
});

app.post('/upload', upload.single('file'), (req, res) => {
  const file = req.file.path;
  loadStudentsFromExcel(file);
  res.redirect('/');
});

app.post('/download', (req, res) => {
  const filePath = 'students.xlsx';

  fs.readFile(filePath, (err, data) => {
    if (err) {
      console.error(err);
      res.status(500).send('Internal Server Error');
      return;
    }

    res.setHeader('Content-Type', 'application/vnd.ms-excel');
    res.setHeader('Content-Disposition', 'attachment; filename=students.xlsx');
    res.send(data);
  });
});

app.listen(port, () => {
  createWorksheet();
  console.log(`Server is listening on port ${port}`);
});

module.exports = app;
