const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx-style");
const fs = require("fs");
const cors = require("cors");
const path = require("path");
const { spawn } = require("child_process"); // For invoking Python script
const stringSimilarity = require("string-similarity");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json());

// Multer setup for file uploads
const upload = multer({ dest: "uploads/" });

// Load mappings and formatting details
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));

// --- Helper Functions ---

// Apply formatting to Excel cells
function applyFormatting(worksheet, formatting) {
    Object.keys(formatting).forEach((cell) => {
        if (worksheet[cell]) {
            const format = formatting[cell];
            worksheet[cell].s = {
                font: {
                    name: format.fontFamily || "Arial",
                    sz: format.fontSize || 10,
                    bold: format.bold || false,
                },
                alignment: {
                    horizontal: format.horizontalAlignment || "left",
                    vertical: format.verticalAlignment || "center",
                },
                fill: {
                    fgColor: { rgb: rgbToHex(format.backgroundColor || { red: 1, green: 1, blue: 1 }) },
                },
            };
        }
    });
}

// RGB to HEX conversion
function rgbToHex(color) {
    const r = Math.round((color.red || 0) * 255).toString(16).padStart(2, "0");
    const g = Math.round((color.green || 0) * 255).toString(16).padStart(2, "0");
    const b = Math.round((color.blue || 0) * 255).toString(16).padStart(2, "0");
    return r + g + b;
}

// Match headings with 90% similarity
function findMatchingCell(key) {
    const keys = Object.keys(mappings);
    const matches = stringSimilarity.findBestMatch(key, keys);
    const bestMatch = matches.bestMatch;

    if (bestMatch.rating >= 0.9) {
        return mappings[bestMatch.target];
    }
    return null; // No match found
}

// --- Upload Endpoint ---
app.post("/upload", upload.any(), async (req, res) => {
    try {
        console.log("Files uploaded successfully.");

        // Get uploaded files
        const docxFile = req.files.find((file) => file.fieldname === "docxFile");
        const excelFile = req.files.find((file) => file.fieldname === "excelFile");

        if (!docxFile || !excelFile) {
            console.error("Missing required files!");
            return res.status(400).json({ status: "error", message: "Missing required files!" });
        }

        // --- Invoke Python script ---
        console.log("Invoking Python script to parse DOCX...");
        const pythonProcess = spawn("python", ["parse_docx.py", docxFile.path]);

        let output = "";
        let error = "";

        pythonProcess.stdout.on("data", (data) => {
            output += data.toString();
        });

        pythonProcess.stderr.on("data", (data) => {
            error += data.toString();
        });

        pythonProcess.on("close", (code) => {
            if (code !== 0 || error) {
                console.error("Python script execution failed:", error);
                return res.status(500).json({ status: "error", message: "Failed to process DOCX file!" });
            }

            console.log("Python script executed successfully.");

            // Read the generated CSV file
            const csvPath = "parsed_data.csv";
            console.log(`Reading generated CSV file: ${csvPath}`);
            const csvData = fs.readFileSync(csvPath, "utf8");

            const rows = csvData.split("\n").slice(1); // Skip header row
            const keyValuePairs = {};
            rows.forEach((row) => {
                const [key, value] = row.split(",");
                if (key && value) {
                    keyValuePairs[key.trim()] = value.trim();
                }
            });

            // Debug: Show extracted key-value pairs
            console.log("Extracted Data from DOCX (Key-Value Pairs):");
            console.table(keyValuePairs); // Tabular debug output for clarity

            // --- Process Excel ---
            console.log("Processing Excel file...");
            const workbook = XLSX.readFile(excelFile.path);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            // Debug: Data mapping process
            console.log("Starting data mapping to Excel cells...");
            Object.keys(keyValuePairs).forEach((key) => {
                const value = keyValuePairs[key];
                const cell = findMatchingCell(key);

                if (cell) {
                    worksheet[cell] = { v: value };
                    console.log(`Mapped -> Key: "${key}" | Value: "${value}" | Cell: "${cell}"`);
                } else {
                    console.warn(`Skipped -> Key: "${key}" (No mapping found)`);
                }
            });

            // Apply formatting
            console.log("Applying formatting to Excel...");
            applyFormatting(worksheet, formattingDetails);

            // Save updated Excel file
            const outputFilePath = `uploads/updated_${Date.now()}.xlsx`;
            XLSX.writeFile(workbook, outputFilePath);
            console.log(`Updated Excel saved at: ${outputFilePath}`);

            // Response
            res.json({
                status: "success",
                downloadExcel: `https://mapstosheetsackend-1.onrender.com/download/${path.basename(outputFilePath)}`,
                downloadCsv: `https://mapstosheetsackend-1.onrender.com/download/${csvPath}`,
            });
        });
    } catch (error) {
        console.error("Error:", error);
        res.json({ status: "error", message: error.toString() });
    }
});

// --- Download Route ---
app.get("/download/:filename", (req, res) => {
    const filename = req.params.filename;
    const filePath = path.join(__dirname, "uploads", filename);
    console.log(`Download requested for file: ${filename}`);
    res.download(filePath, (err) => {
        if (err) {
            console.error("Download error:", err);
            res.status(500).send("Error downloading the file.");
        }
    });
});

// --- Start Server ---
app.listen(5000, () => {
    console.log("Server running on http://localhost:5000");
});
