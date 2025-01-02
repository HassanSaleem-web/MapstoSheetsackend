const express = require("express");
const multer = require("multer");
const pdfParse = require("pdf-parse");
const XLSX = require("xlsx-style"); // For formatting support
const fs = require("fs");
const cors = require("cors");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json());

// Multer setup for file uploads
const upload = multer({ dest: "uploads/" });

// Load mappings and formatting details
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));

// ---- Helper Functions ----

// Parse PDF and start recording from "AGE GROUP"
async function parsePdf(filePath) {
  console.log(`Parsing PDF: ${filePath}`);
  const dataBuffer = fs.readFileSync(filePath);
  const data = await pdfParse(dataBuffer);

  // Split content into lines
  const lines = data.text.split("\n").map((line) => line.trim());
  let startRecording = false; // Flag to detect where to start recording
  const results = {};
  let pairs = [];
  let key = null;

  console.log("Extracting key-value pairs...");
  for (let i = 0; i < lines.length; i++) {
    const line = lines[i];

    // Start recording when "AGE GROUP" is found
    if (line.toUpperCase() === "AGE GROUP") {
      startRecording = true;
    }

    if (!startRecording) continue; // Skip until "AGE GROUP" is found

    // Skip empty lines or ignore headings
    if (line === "" || line.toUpperCase() === "QUESTIONNAIRE" || line.toUpperCase() === "QUESTIONS") {
      continue;
    }

    if (line === line.toUpperCase()) {
      // Treat uppercase lines as headings (keys)
      key = line;
    } else if (key) {
      // Next line is the value
      results[key] = line; // Create key-value pair
      pairs.push(`${key}: ${line}`); // Save pair for txt file
      key = null; // Reset key
    }
  }

  // Write results to a txt file
  const txtFilePath = `uploads/results_${Date.now()}.txt`;
  fs.writeFileSync(txtFilePath, pairs.join("\n"));
  console.log(`Key-value pairs saved in: ${txtFilePath}`);

  return { results, txtFilePath };
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

// ---- Upload Route ----
app.post("/upload", upload.any(), async (req, res) => {
  try {
    console.log("Files uploaded successfully.");

    // Extract uploaded files
    const pdfFile = req.files.find((file) => file.fieldname === "pdfFile");
    const excelFile = req.files.find((file) => file.fieldname === "excelFile");

    if (!pdfFile || !excelFile) {
      return res.status(400).json({ status: "error", message: "Missing required files!" });
    }

    // Parse PDF and extract key-value pairs
    const { results, txtFilePath } = await parsePdf(pdfFile.path);

    // Load Excel sheet
    const excelFilePath = excelFile.path;
    const workbook = XLSX.readFile(excelFilePath);
    const sheetName = workbook.SheetNames[0];
    const worksheet = workbook.Sheets[sheetName];
    console.log(`Loaded Excel Sheet: ${sheetName}`);

    // Map extracted values to Excel cells
    console.log("Mapping values to Excel cells...");
    for (const key in results) {
      if (mappings[key]) {
        const cell = mappings[key]; // Get cell address from mappings.json
        console.log(`Mapping: ${key} -> ${results[key]} -> Cell: ${cell}`);
        worksheet[cell] = { v: results[key] };
      } else {
        console.warn(`No mapping found for key: ${key}`);
      }
    }

    // Apply formatting
    console.log("Applying formatting...");
    applyFormattingToSheet(worksheet, formattingDetails);

    // Set column widths
    console.log("Setting column widths...");
    setColumnWidths(worksheet);

    // Save updated Excel file
    const outputFilePath = `uploads/updated_${Date.now()}.xlsx`;
    XLSX.writeFile(workbook, outputFilePath);
    console.log(`Updated Excel file saved at: ${outputFilePath}`);

    // Respond with download links
    res.json({
      status: "success",
      downloadExcel: `https://mapstosheetsackend-1.onrender.com/download/${outputFilePath.split("/").pop()}`,
      downloadTxt: `http://localhost:5000/download/${txtFilePath.split("/").pop()}`,
    });
  } catch (error) {
    console.error("Error:", error);
    res.json({ status: "error", message: error.toString() });
  }
});

// ---- Download Route ----
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

// Start Server
app.listen(5000, () => {
  console.log("Server running on http://localhost:5000");
});
