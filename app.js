const express = require('express');
const multer = require('multer');
const ExcelJS = require('exceljs');
const fs = require('fs');
const path = require('path');
const cors = require('cors');

const app = express();
app.use(cors());

const upload = multer({ dest: 'uploads/' });
const scriptDir = __dirname;

app.get('/templates', (req, res) => {
    const templateFiles = fs.readdirSync(__dirname).filter(file => file.endsWith('.xlsx'));
    res.json(templateFiles);
});

app.post('/process', upload.single('uploadFile'), async (req, res) => {
    try {
        const uploadedFilePath = req.file.path;
        const selectedTemplate = req.body.template;

        if (!selectedTemplate || !selectedTemplate.endsWith('.xlsx')) {
            throw new Error('Invalid selected template.');
        }

        const selectedTemplatePath = path.join(__dirname, selectedTemplate);

        // Read the uploaded file
        const uploadedWorkbook = new ExcelJS.Workbook();
        await uploadedWorkbook.xlsx.readFile(uploadedFilePath);
        console.log('Uploaded workbook read successfully.');

        // Find "Input" sheet in uploaded file
        let inputSheet;
        uploadedWorkbook.eachSheet(sheet => {
            if (sheet.name === 'Input') {
                inputSheet = sheet;
                return false; // Stop iteration
            }
        });
        if (!inputSheet) {
            throw new Error('No "Input" sheet found in the uploaded file.');
        }

        // Read the selected template file
        const templateWorkbook = new ExcelJS.Workbook();
        await templateWorkbook.xlsx.readFile(selectedTemplatePath);

        // Remove existing "Input" sheet in template file, if any
        templateWorkbook.eachSheet((sheet, sheetId) => {
            if (sheet.name === 'Input') {
                templateWorkbook.removeWorksheet(sheetId);
            }
        });

        // Add "Input" sheet and copy data from uploaded file to template file
        const newInputSheet = templateWorkbook.addWorksheet('Input', {
            properties: inputSheet.properties
        });
        inputSheet.eachRow((row, rowNumber) => {
            row.eachCell((cell, colNumber) => {
                const newCell = newInputSheet.getCell(rowNumber, colNumber);
                newCell.value = cell.result !== undefined ? cell.result : cell.value; // Use the cell's result if it's a formula, otherwise use its value
                newCell.style = cell.style;
            });
        });

        // Process the "Output" sheet in the template file
        let outputSheet = templateWorkbook.getWorksheet('Output');
        if (outputSheet) {
            outputSheet.eachRow((row, rowNumber) => {
                row.eachCell((cell, colNumber) => {
                    if (cell.type === ExcelJS.ValueType.Formula) {
                        cell.value = cell.result; // Replace the formula with its result
                    }
                });
            });
        }

        // Remove all sheets except "Output"
        templateWorkbook.eachSheet((sheet, sheetId) => {
            if (sheet.name !== 'Output') {
                templateWorkbook.removeWorksheet(sheetId);
            }
        });

        // Convert workbook to buffer
        const buffer = await templateWorkbook.xlsx.writeBuffer();
        console.log('Workbook processed successfully.');

        // Clean up uploaded file
        fs.unlinkSync(uploadedFilePath);
        console.log('Uploaded file cleaned up.');

        // Send the buffer as response
        res.setHeader('Content-Disposition', 'attachment; filename=Processed.xlsx');
        res.send(buffer);
    } catch (err) {
        console.error('Error processing files:', err);
        res.status(500).send('Error processing files: ' + err.message);
    }
});
// Serve HTML form for uploading new templates


// Route for uploading new templates
app.post("/template", upload.single("templateFile"), (req, res) => {
    const templateFilePath = req.file ? req.file.path : null;

    if (!templateFilePath) {
        return res.status(400).send("Template file not provided.");
    }

    const originalFileName = req.file.originalname;
    const newFilePath = path.join(scriptDir, originalFileName);

    fs.rename(templateFilePath, newFilePath, (err) => {
        if (err) {
            console.error("Error moving uploaded file:", err);
            return res.status(500).send("Error uploading template file.");
        }
        res.send("Template file uploaded successfully.");
    });
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
    console.log(`Server is running on port ${PORT}`);
});
