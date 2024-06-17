const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const bodyParser = require("body-parser");

const app = express();
const port = 3000;

app.use(bodyParser.urlencoded({ extended: true }));

const upload = multer({ dest: "uploads/" });

// Serve HTML form
app.get("/", (req, res) => {
  const formHtml = `
    <!DOCTYPE html>
    <html lang="en">
    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>Upload Excel File</title>
    </head>
    <body>
      <h1>Upload Excel File and Specify Sheet Name</h1>
      <form action="/upload" method="post" enctype="multipart/form-data">
        <label for="file">Select Excel File:</label><br>
        <input type="file" id="file" name="file" accept=".xlsx"><br><br>
        <label for="sheetName">Sheet Name:</label><br>
        <input type="text" id="sheetName" name="sheetName"><br><br>
        <button type="submit">Upload and Process</button>
      </form>
    </body>
    </html>
  `;
  res.send(formHtml);
});

// Handle file upload and processing
app.post("/upload", upload.single("file"), async (req, res) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }

  const sheetName = req.body.sheetName;
  const filePath = req.file.path;

  try {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);

    // Delete all sheets except the specified sheetName
    workbook.eachSheet((sheet, sheetId) => {
      if (sheet.name !== sheetName) {
        workbook.removeWorksheet(sheetId);
      } else {
        // Ensure all cell values are static
        sheet.eachRow((row) => {
          row.eachCell({ includeEmpty: true }, (cell) => {
            if (cell.type === ExcelJS.ValueType.Formula) {
              // Convert formulas to static values
              cell.value = cell.result || 0; // Use result if available, else default to 0
              cell.type = ExcelJS.ValueType.Number; // Change type to number (or string if necessary)
            }
            // Handle localization or non-English characters if needed
            // Example: Convert non-ASCII characters to base values
            if (
              typeof cell.value === "string" &&
              /[^\x00-\x7F]/.test(cell.value)
            ) {
              cell.value = convertToBaseValue(cell.value); // Implement your conversion logic
            }
          });
        });
      }
    });

    const tempOutputFilePath = path.join("uploads", `${sheetName}.xlsx`);

    // Example: Ensure gridlines are set
    const sheet = workbook.getWorksheet(sheetName);
    sheet.views = [
      {
        showGridLines: false, // Ensure gridlines are visible in the output
      },
    ];

    await workbook.xlsx.writeFile(tempOutputFilePath);
    fs.unlinkSync(filePath);

    res.download(tempOutputFilePath, `${sheetName}.xlsx`, (err) => {
      if (err) {
        console.error(err);
        res.status(500).send("An error occurred during file download.");
      } else {
        fs.unlinkSync(tempOutputFilePath); // Clean up the temporary file
      }
    });
  } catch (error) {
    fs.unlinkSync(filePath);
    console.error(error);
    res.status(500).send("An error occurred during file processing.");
  }
});

// Start the server
app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});

// Function to convert non-ASCII characters to base values
function convertToBaseValue(str) {
  // Implement your conversion logic here
  // Example: Replace non-ASCII characters with their ASCII equivalents
  return str.replace(/[^\x00-\x7F]/g, "");
}
