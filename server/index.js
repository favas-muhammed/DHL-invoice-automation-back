const express = require("express");
const mongoose = require("mongoose");
const errorHandler = require("../middleware/errorHandler");
const multer = require("multer");
const XLSX = require("xlsx");
const pdfParse = require("pdf-parse");
const cors = require("cors");
require("dotenv").config();

const app = express();
const port = process.env.PORT || 5000;

app.use(
  cors({
    origin: "http://localhost:5173", // Allow requests from this origin
    methods: ["GET", "POST"], // Specify allowed methods
    credentials: true, // Allow credentials if needed
  })
);

app.use(express.json());

// Configure multer for PDF uploads
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 }, // 10MB limit
});

// Helper function to extract text from PDF
async function extractTextFromPDF(pdfBuffer) {
  try {
    const data = await pdfParse(pdfBuffer);
    return data.text;
  } catch (error) {
    throw new Error("Failed to extract text from PDF: " + error.message);
  }
}
function formatAmount(amount) {
  // Convert amount to a number and format it with commas
  const number = parseFloat(amount);
  return number.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}
// Helper function to find data in PDF text
function findDataInPDF(pdfText, searchTerm) {
  const lines = pdfText.split("\n");
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes(searchTerm)) {
      // Look for the last amount value before the next ID or end of section
      let j = i;
      let lastAmount = null;
      let previousAmount = null;
      while (j < lines.length) {
        // Stop if we hit another ID or end of section
        if (
          j !== i &&
          (lines[j].match(/\d{10}/) || lines[j].includes("Total facture"))
        ) {
          break;
        }
        // Look for amount pattern (numbers with optional comma and 2 decimal places)
        const amountMatch = lines[j].match(
          /(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})\s*$/
        );
        if (amountMatch) {
          // Store the previous amount before updating lastAmount
          previousAmount = lastAmount;
          // Remove commas and convert to number
          lastAmount = amountMatch[1].replace(/,/g, "");
        }
        // If we hit a Sous-Total Service line, use the previous amount instead
        if (lines[j].startsWith('Sous-Total Service') && previousAmount) {
          return formatAmount(previousAmount);
        }
        j++;
      }
      return lastAmount ? formatAmount(lastAmount) : null;
    }
  }

  return null;
}

// Helper function to process Excel and add separator rows
async function processExcelWithSeparators(buffer, referenceColumn) {
  const workbook = XLSX.read(buffer);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  // Convert column letter to index (e.g., 'A' -> 0, 'B' -> 1, 'AA' -> 26)
  const colIndex = referenceColumn.split('').reduce((acc, char) => 
    acc * 26 + char.toUpperCase().charCodeAt(0) - 'A'.charCodeAt(0) + 1, 0) - 1;

  const newRows = [];
  let currentGroup = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const value = row[colIndex];

    if (currentGroup !== value) {
      // Add two empty rows between groups (except at the start)
      if (currentGroup !== null) {
        newRows.push(Array(row.length).fill(""));
        newRows.push(Array(row.length).fill(""));
      }
      currentGroup = value;
    }
    newRows.push(row);
  }

  // Create new workbook with processed rows
  const newWorkbook = XLSX.utils.book_new();
  const newWorksheet = XLSX.utils.aoa_to_sheet(newRows);
  XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

  return XLSX.write(newWorkbook, { type: "buffer", bookType: "xlsx" });
}

app.post("/process-excel", upload.single("excel"), async (req, res) => {
  try {
    if (!req.file || !req.body.referenceColumn) {
      throw new Error("Excel file and reference column are required");
    }

    const processedBuffer = await processExcelWithSeparators(
      req.file.buffer,
      req.body.referenceColumn
    );

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader(
      "Content-Disposition",
      "attachment; filename=processed_excel.xlsx"
    );
    res.send(processedBuffer);
  } catch (error) {
    console.error("Error processing Excel:", error);
    res.status(400).json({
      error: error.message,
      success: false,
    });
  }
});

app.post("/upload",
  upload.fields([
    { name: "pdfs", maxCount: 100 }, // Allow up to 100 PDFs
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      if (!req.files?.pdfs || !req.files?.excel) {
        const error = new Error(
          "At least one PDF and one Excel file are required"
        );
        error.statusCode = 400;
        throw error;
      }

      // 1. Extract text from all PDFs
      let combinedPdfText = "";
      for (const pdfFile of req.files.pdfs) {
        const pdfText = await extractTextFromPDF(pdfFile.buffer);
        combinedPdfText += pdfText + "\n";
      }

      // 2. Read Excel file
      const workbook = XLSX.read(req.files.excel[0].buffer);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      const updates = [];

      // 3. Process each row (skip header)
      for (let i = 1; i < rows.length; i++) {
        const searchTerm = rows[i][0]; // First column
        if (searchTerm) {
          const foundData = findDataInPDF(combinedPdfText, searchTerm);
          if (foundData) {
            rows[i][1] = foundData; // Update second column
            updates.push([searchTerm, foundData]);
          }
        }
      }

      // 4. Write back to Excel
      const newWorkbook = XLSX.utils.book_new();
      const newWorksheet = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

      // 5. Generate Excel file
      const excelBuffer = XLSX.write(newWorkbook, {
        type: "buffer",
        bookType: "xlsx",
      });

      // 6. Send the updated Excel file back to client
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      res.setHeader(
        "Content-Disposition",
        "attachment; filename=updated_excel.xlsx"
      );
      res.send(excelBuffer);
    } catch (error) {
      console.error("Error processing upload:", error);
      const statusCode = error.statusCode || 500;
      let message = "Internal server error while processing upload";

      // Handle specific error cases
      if (error.code === "LIMIT_FILE_SIZE") {
        message = "File size too large. Maximum size is 10MB";
      } else if (error.message.includes("Failed to extract text from PDF")) {
        message = "Invalid or corrupted PDF file";
      } else if (error.message) {
        message = error.message;
      }

      res.status(statusCode).json({
        error: message,
        success: false,
      });
    }
  }
);

// Error handling middleware should be last
app.use(errorHandler);

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

// Handle unhandled promise rejections
process.on("unhandledRejection", (err) => {
  console.error("Unhandled Promise Rejection:", err);
  process.exit(1);
});