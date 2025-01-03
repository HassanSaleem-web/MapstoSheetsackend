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
console.log("Loading mappings.json and formatting_details.json...");
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));
console.log("Mappings loaded:", mappings);
console.log("Formatting loaded:", formattingDetails);

// --- Helper Functions ---
// Apply formatting to Excel cells
function applyFormatting(worksheet, formatting) {
    console.log("Applying formatting to worksheet...");
    Object.keys(formatting).forEach((cell) => {
        if (worksheet[cell]) {
            const format = formatting[cell];
            console.log(`Applying formatting to cell ${cell}:`, format);
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
        } else {
            console.warn(`No formatting found for cell: ${cell}`);
        }
    });
    console.log("Formatting applied successfully.");
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
    console.log(`Finding match for heading: "${key}"`);
    const keys = Object.keys(mappings);
    const matches = stringSimilarity.findBestMatch(key, keys);
    const bestMatch = matches.bestMatch;

    console.log(`Best match for "${key}": "${bestMatch.target}" with similarity ${bestMatch.rating}`);
    if (bestMatch.rating >= 0.9) {
        return mappings[bestMatch.target];
    }
    console.warn(`No match found for "${key}"`);
    return null; // No match found
}

// --- Upload Endpoint ---
app.post("/upload", upload.any(), async (req, res) => {
    try {
        console.log("Files uploaded successfully.");
        console.log("Uploaded Files:", req.files);

        const docxFile = req.files.find((file) => file.fieldname === "docxFile");
        const excelFile = req.files.find((file) => file.fieldname === "excelFile");

        if (!docxFile || !excelFile) {
            console.error("Missing required files!");
            return res.status(400).json({ status: "error", message: "Missing required files!" });
        }

        // --- Invoke Python script as a child process ---
        console.log("Invoking Python script to parse DOCX...");
        const pythonProcess = spawn("python", ["parse_docx.py", docxFile.path]);

        let output = "";
        let error = "";

        pythonProcess.stdout.on("data", (data) => {
            output += data.toString();
            console.log("Python Output:", data.toString());
        });

        pythonProcess.stderr.on("data", (data) => {
            error += data.toString();
            console.error("Python Error:", data.toString());
        });

        pythonProcess.on("close", (code) => {
            if (code !== 0 || error) {
                console.error("Python script execution failed:", error);
                return res.status(500).json({ status: "error", message: "Failed to process DOCX file!" });
            }

            console.log("Python script executed successfully.");

            // Read the generated CSV file
            const csvPath = "parsed_data.csv"; // Hardcoded in Python script
            console.log(`Reading generated CSV file: ${csvPath}`);
            const csvData = fs.readFileSync(csvPath, "utf8");
            console.log("CSV Data:\n", csvData);

            const rows = csvData.split("\n").slice(1); // Skip header row
            const keyValuePairs = {};
            rows.forEach((row) => {
                const [key, value] = row.split(",");
                if (key && value) {
                    keyValuePairs[key.trim()] = value.trim();
                }
            });
            console.log("Extracted Key-Value Pairs:", keyValuePairs);

            // --- Process Excel ---
            console.log("Processing Excel file...");
            const workbook = XLSX.readFile(excelFile.path);
            const worksheet = workbook.Sheets[workbook.SheetNames[0]];

            Object.keys(keyValuePairs).forEach((key) => {
                const value = keyValuePairs[key];
                const cell = findMatchingCell(key);

                if (cell) {
                    console.log(`Mapping: ${key} -> ${value} -> ${cell}`);
                    worksheet[cell] = { v: value };
                } else {
                    console.warn(`No mapping found for: ${key}`);
                }
            });

            // Apply formatting
            applyFormatting(worksheet, formattingDetails);

            // Save updated Excel file
            const outputFilePath = `uploads/updated_${Date.now()}.xlsx`;
            XLSX.writeFile(workbook, outputFilePath);
            console.log(`Updated Excel file saved at: ${outputFilePath}`);

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
    console.log(`Downloading file: ${filename} from path: ${filePath}`);
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
