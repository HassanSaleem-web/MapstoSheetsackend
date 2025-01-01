const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx-style"); // Supports cell formatting
const fs = require("fs");
const cors = require("cors");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json());

// Multer setup for file uploads (accepts any field name)
const upload = multer({ dest: "uploads/" });

// Load mappings and formatting details
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));

// Parse PDF (Extract headings and values)
async function parsePdf(filePath) {
  console.log(`Parsing PDF: ${filePath}`);
  const dataBuffer = fs.readFileSync(filePath);
  const data = await pdfParse(dataBuffer);
  console.log("Extracted PDF text:", data.text.substring(0, 100)); // Debug log
  return data.text;
}

// Apply formatting dynamically from formatting_details.json
function applyFormattingToSheet(worksheet, formatting) {
  Object.keys(formatting).forEach((cell) => {
    if (worksheet[cell]) {
      const format = formatting[cell];

      // Apply styles
      worksheet[cell].s = {
        font: {
          name: format.fontFamily || "Arial",
          sz: format.fontSize || 10,
          bold: format.bold || false,
          italic: format.italic || false,
          underline: format.underline || false,
        },
        alignment: {
          horizontal: format.horizontalAlignment ? format.horizontalAlignment.toLowerCase() : "left",
          vertical: format.verticalAlignment ? format.verticalAlignment.toLowerCase() : "center",
        },
        fill: {
          fgColor: {
            rgb: rgbToHex(format.backgroundColor || { red: 1, green: 1, blue: 1 }),
          },
        },
      };
    } else {
      console.warn(`No formatting details found for cell: ${cell}`);
    }
  });
}

// Convert RGB to HEX for Excel formatting
function rgbToHex(color) {
  const r = Math.round((color.red || 0) * 255).toString(16).padStart(2, "0");
  const g = Math.round((color.green || 0) * 255).toString(16).padStart(2, "0");
  const b = Math.round((color.blue || 0) * 255).toString(16).padStart(2, "0");
  return r + g + b;
}

// Set column widths
function setColumnWidths(worksheet) {
  worksheet['!cols'] = [
    { wch: 30 }, // Column A width
    { wch: 50 }, // Column B width
    { wch: 40 }, // Column C width
  ];
}

// Handle uploads and process data
app.post(
  "/upload",
  upload.any(), // Accepts any field name
  async (req, res) => {
    try {
      console.log("Files uploaded successfully.");

      // Debugging: Log incoming fields and files
      console.log("Incoming Fields:", req.body);
      console.log("Incoming Files:", req.files);

      // Extract uploaded files
      const pdfFile = req.files.find((file) => file.fieldname === "pdfFile");
      const excelFile = req.files.find((file) => file.fieldname === "excelFile");

      // Validate uploaded files
      if (!pdfFile || !excelFile) {
        return res.status(400).json({ status: "error", message: "Missing required files!" });
      }

      // Parse uploaded PDF
      const pdfText = await parsePdf(pdfFile.path);
      const excelFilePath = excelFile.path;

      // Extract data from PDF
      const lines = pdfText.split("\n");
      const data = {};
      console.log("Extracting key-value pairs from PDF...");

      for (let i = 0; i < lines.length; i++) {
        if (lines[i].trim().toUpperCase() && i + 1 < lines.length) {
          const key = lines[i].trim().toUpperCase();
          const value = lines[i + 1].trim();
          data[key] = value;
          console.log(`Extracted: ${key} => ${value}`);
          i++;
        }
      }

      // Load Excel sheet
      const workbook = XLSX.readFile(excelFilePath);
      const sheetName = workbook.SheetNames[0]; // Assume first sheet
      const worksheet = workbook.Sheets[sheetName];
      console.log(`Loaded Excel Sheet: ${sheetName}`);

      // Match headings and map values
      console.log("Matching headings and mapping values...");
      for (const key in data) {
        if (mappings[key]) {
          const cell = mappings[key]; // Get cell address from mappings.json
          console.log(`Mapping: ${key} -> ${data[key]} -> Cell: ${cell}`);
          worksheet[cell] = { v: data[key] };
          console.log(`Updated Cell ${cell} with Value: ${data[key]}`);
        } else {
          console.warn(`No mapping found for key: ${key}`);
        }
      }

      // Apply formatting dynamically
      console.log("Applying formatting dynamically...");
      applyFormattingToSheet(worksheet, formattingDetails);

      // Set column widths
      console.log("Setting column widths...");
      setColumnWidths(worksheet);

      // Save updated Excel file
      const outputFilePath = `uploads/updated_${Date.now()}.xlsx`;
      XLSX.writeFile(workbook, outputFilePath);
      console.log(`Updated Excel file saved at: ${outputFilePath}`);

      res.json({
        status: "success",
        downloadLink: `https://mapstosheetsackend-1.onrender.com/download/${outputFilePath.split("/").pop()}`,
      });
    } catch (error) {
      console.error("Error:", error);
      res.json({ status: "error", message: error.toString() });
    }
  }
);

// Serve updated Excel files for download
app.get("/download/:filename", (req, res) => {
  const filename = req.params.filename;
  const filePath = `uploads/${filename}`;
  res.download(filePath, (err) => {
    if (err) {
      console.error("Download error:", err);
      res.status(500).send("Error downloading the file.");
    }
  });
});

// Start server
app.listen(5000, () => {
  console.log("Server running on http://localhost:5000");
});
