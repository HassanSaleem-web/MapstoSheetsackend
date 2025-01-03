const express = require("express");
const multer = require("multer");
const XLSX = require("xlsx-style");
const fs = require("fs");
const cors = require("cors");
const path = require("path");
const docxParser = require("docx-parser"); // New package for DOCX parsing
const stringSimilarity = require("string-similarity");
require("dotenv").config();

const app = express();
app.use(cors());
app.use(express.json());

// Serve static files for uploads
app.use("/uploads", express.static(path.join(__dirname, "uploads")));

// Multer setup for file uploads
const upload = multer({ dest: "uploads/" });

// Load mappings and formatting details
const mappings = JSON.parse(fs.readFileSync("mappings.json", "utf8"));
const formattingDetails = JSON.parse(fs.readFileSync("formatting_details.json", "utf8"));

// --- Helper Functions ---

// Parse DOCX and extract key-value pairs
function parseDocx(docxPath) {
    console.log(`Parsing DOCX file: ${docxPath}`);

    return new Promise((resolve, reject) => {
        docxParser.parseDocx(docxPath, (err, data) => {
            if (err) {
                reject(err); // Handle errors
            } else {
                const keyValuePairs = {};
                let currentKey = "";

                // Split text into lines and analyze each line
                const lines = data.split("\n");
                lines.forEach((line) => {
                    const trimmed = line.trim();
                    if (trimmed) {
                        // Treat uppercase text as headings
                        if (trimmed === trimmed.toUpperCase()) {
                            if (currentKey) {
                                keyValuePairs[currentKey] =
                                    keyValuePairs[currentKey].trim() || "No response";
                            }
                            currentKey = trimmed; // Set new heading
                            keyValuePairs[currentKey] = ""; // Initialize value
                        } else if (currentKey) {
                            // Append non-uppercase text as values
                            keyValuePairs[currentKey] += " " + trimmed;
                        }
                    }
                });

                // Save the last key-value pair
                if (currentKey) {
                    keyValuePairs[currentKey] =
                        keyValuePairs[currentKey].trim() || "No response";
                }

                console.log("Extracted Key-Value Pairs:", keyValuePairs); // Debug log
                resolve(keyValuePairs); // Return parsed data
            }
        });
    });
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

        // Extract uploaded files
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
            downloadExcel: `/uploads/${path.basename(outputFilePath)}`,
            downloadCsv: `/uploads/${path.basename(csvPath)}`,
        });
    } catch (error) {
        console.error("Error:", error);
        res.json({ status: "error", message: error.toString() });
    }
});

// --- Start Server ---
app.listen(5000, () => {
    console.log("Server running on http://localhost:5000");
});
