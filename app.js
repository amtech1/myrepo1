const express = require("express");
const multer = require("multer");
const ExcelJS = require("exceljs");
const fs = require("fs");
const path = require("path");
const bodyParser = require("body-parser");

const app = express();
const port = 3000;
const scriptDir = __dirname; // Directory where the script is located
const upload = multer({ dest: "uploads/" }); // Upload directory for the new data files

app.use(bodyParser.urlencoded({ extended: true }));

// Serve HTML form
app.get("/", (req, res) => {
  fs.readdir(scriptDir, (err, files) => {
    if (err) {
      return res.status(500).send("Unable to scan files.");
    }

    const xlsxFiles = files.filter((file) => path.extname(file) === ".xlsx");

    let fileOptions = xlsxFiles
      .map((file) => `<option value="${file}">${file}</option>`)
      .join("");

    const formHtml = `
      <!DOCTYPE html>
      <html lang="en">
      <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Select Template File</title>
      </head>
      <body>
        <h1>Select Template File and Upload Data File</h1>
        <form action="/process" method="post" enctype="multipart/form-data">
          <label for="templateFileName">Choose Template File:</label><br>
          <select id="templateFileName" name="templateFileName">
            ${fileOptions}
          </select><br><br>
          <label for="dataFile">Upload Data File:</label><br>
          <input type="file" id="dataFile" name="dataFile" accept=".xlsx"><br><br>
          <button type="submit">Upload and Process</button>
        </form>
      </body>
      </html>
    `;
    res.send(formHtml);
  });
});

// Handle file selection and processing
app.post("/process", upload.single("dataFile"), async (req, res) => {
  const templateFileName = req.body.templateFileName;
  const dataFilePath = req.file ? req.file.path : null;

  if (!templateFileName || !dataFilePath) {
    return res.status(400).send("Template file or data file not provided.");
  }

  const templateFilePath = path.join(scriptDir, templateFileName);
  const sheetName = "Output";

  try {
    // Load the template workbook
    const templateWorkbook = new ExcelJS.Workbook();
    await templateWorkbook.xlsx.readFile(templateFilePath);

    // Load the data workbook
    const dataWorkbook = new ExcelJS.Workbook();
    await dataWorkbook.xlsx.readFile(dataFilePath);

    // Get the first sheet from the data workbook
    const dataSheet = dataWorkbook.worksheets[0];

    // Get the input sheet from the template workbook
    const inputSheet = templateWorkbook.getWorksheet("Input");
    if (!inputSheet) {
      throw new Error("Input sheet not found in the template file.");
    }

    // Clear existing data in the input sheet (optional)
    inputSheet.eachRow((row) => {
      row.eachCell((cell) => {
        cell.value = null;
      });
    });

    // Copy data from the data sheet to the input sheet
    dataSheet.eachRow((row, rowNumber) => {
      row.eachCell((cell, colNumber) => {
        inputSheet.getCell(rowNumber, colNumber).value = cell.value;
      });
    });

    // Ensure all cell values in the output sheet are static
    const outputSheet = templateWorkbook.getWorksheet(sheetName);
    if (!outputSheet) {
      throw new Error("Output sheet not found in the template file.");
    }

    outputSheet.eachRow((row) => {
      row.eachCell({ includeEmpty: true }, (cell) => {
        if (cell.type === ExcelJS.ValueType.Formula) {
          cell.value = cell.result || 0; // Use result if available, else default to 0
          cell.type = ExcelJS.ValueType.Number; // Change type to number (or string if necessary)
        }
        // Handle localization or non-English characters if needed
        if (typeof cell.value === "string" && /[^\x00-\x7F]/.test(cell.value)) {
          cell.value = convertToBaseValue(cell.value); // Implement your conversion logic
        }
      });
    });

    // Remove all sheets except the output sheet
    templateWorkbook.eachSheet((sheet, sheetId) => {
      if (sheet.name !== sheetName) {
        templateWorkbook.removeWorksheet(sheetId);
      }
    });

    const tempOutputFilePath = path.join(
      scriptDir,
      `processed_${templateFileName}`
    );

    // Preserve the original gridline settings
    if (outputSheet) {
      outputSheet.views = outputSheet.views.map((view) => ({
        ...view,
        showGridLines: view.showGridLines,
      }));
    }

    await templateWorkbook.xlsx.writeFile(tempOutputFilePath);

    res.download(tempOutputFilePath, `processed_${templateFileName}`, (err) => {
      if (err) {
        console.error(err);
        res.status(500).send("An error occurred during file download.");
      } else {
        fs.unlinkSync(tempOutputFilePath); // Clean up the temporary file
        fs.unlinkSync(dataFilePath); // Clean up the uploaded data file
      }
    });
  } catch (error) {
    console.error(error);
    res.status(500).send("An error occurred during file processing.");
    if (dataFilePath) {
      fs.unlinkSync(dataFilePath); // Clean up the uploaded data file
    }
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
