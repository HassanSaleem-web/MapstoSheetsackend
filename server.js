const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx-style");
const fs = require("fs");
const cors = require("cors");
const path = require("path");
const mammoth = require("mammoth");
const stringSimilarity = require("string-similarity");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json());

// Serve static files for uploads
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

// Multer setup for file uploads
const upload = multer({ dest: "uploads/" });

// Load mappings and formatting
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));

// --- Helper Functions ---

// Parse DOCX and extract key-value pairs
async function parseDocx(docxPath) {
    console.log(`Parsing DOCX file: ${docxPath}`);
    const result = await mammoth.extractRawText({ path: docxPath });
    const text = result.value; // Extract text from the document

    const keyValuePairs = {};
    let currentKey = "";

    const lines = text.split("\n");
    lines.forEach((line) => {
        const trimmed = line.trim();
        if (trimmed) {
            if (trimmed === trimmed.toUpperCase()) { // Heading (uppercase or bold-like behavior)
                currentKey = trimmed;
                keyValuePairs[currentKey] = "";
            } else if (currentKey) {
                keyValuePairs[currentKey] += " " + trimmed; // Append multi-line values
            }
        }
    });

    console.log("Parsed Key-Value Pairs:", keyValuePairs); // Debug log
    return keyValuePairs;
}

// Save key-value pairs to CSV
function saveToCsv(data, outputCsv) {
    console.log(`Saving parsed data to CSV: ${outputCsv}`);
    const rows = [["Heading", "Value"]];
    for (const [key, value] of Object.entries(data)) {
        rows.push([key, value]);
    }
    const csvContent = rows.map((row) => row.join(",")).join("\n");
    fs.writeFileSync(outputCsv, csvContent);
}

// Match headings with 90% similarity
function findMatchingCell(key) {
    const keys = Object.keys(mappings);
    const matches = stringSimilarity.findBestMatch(key, keys);
    const bestMatch = matches.bestMatch;

    console.log(`Matching key: ${key}, Best Match: ${bestMatch.target}, Similarity: ${bestMatch.rating}`);
    if (bestMatch.rating >= 0.9) {
        return mappings[bestMatch.target];
    }
    console.warn(`No mapping found for key: ${key}`);
    return null; // No match found
}

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

// RGB to HEX
function rgbToHex(color) {
    const r = Math.round((color.red || 0) * 255).toString(16).padStart(2, "0");
    const g = Math.round((color.green || 0) * 255).toString(16).padStart(2, "0");
    const b = Math.round((color.blue || 0) * 255).toString(16).padStart(2, "0");
    return r + g + b;
}

// --- Upload Endpoint ---
app.post("/upload", upload.any(), async (req, res) => {
    try {
        console.log("Files uploaded successfully.");

        const docxFile = req.files.find((file) => file.fieldname === "docxFile");
        const excelFile = req.files.find((file) => file.fieldname === "excelFile");

        if (!docxFile || !excelFile) {
            return res.status(400).json({ status: "error", message: "Missing required files!" });
        }

        // Parse DOCX and extract key-value pairs
        const keyValuePairs = await parseDocx(docxFile.path);
        const csvPath = `uploads/results_${Date.now()}.csv`;
        saveToCsv(keyValuePairs, csvPath);

        // Load Excel template
        const workbook = XLSX.readFile(excelFile.path);
        const worksheet = workbook.Sheets[workbook.SheetNames[0]];

        // Map values to Excel cells
        console.log("Mapping values to Excel cells...");
        Object.keys(keyValuePairs).forEach((key) => {
            const value = keyValuePairs[key];
            const cell = findMatchingCell(key);

            if (cell) {
                worksheet[cell] = { v: value };
                console.log(`Mapped: ${key} -> ${value} -> ${cell}`);
            }
        });

        // Apply formatting and save
        console.log("Applying formatting...");
        applyFormatting(worksheet, formattingDetails);

        const outputFilePath = `uploads/updated_${Date.now()}.xlsx`;
        XLSX.writeFile(workbook, outputFilePath);

        console.log(`Excel saved at: ${outputFilePath}`);
        console.log(`CSV saved at: ${csvPath}`);

        res.json({
            status: "success",
            downloadExcel: `https://mapstosheetsackend-1.onrender.com/uploads/${path.basename(outputFilePath)}`,
            downloadCsv: `https://mapstosheetsackend-1.onrender.com/uploads/${path.basename(csvPath)}`,
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

    console.log(`Attempting to download: ${filePath}`); // Debug log

    if (!fs.existsSync(filePath)) {
        console.error("File not found:", filePath); // Debug log
        return res.status(404).send("File not found.");
    }

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
