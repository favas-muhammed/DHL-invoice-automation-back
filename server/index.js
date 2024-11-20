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

// Helper function to find data in PDF text
function findDataInPDF(pdfText, searchTerm) {
  // This is a simple implementation. You might need to adjust based on your PDF structure
  const lines = pdfText.split("\n");
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes(searchTerm)) {
      // Assuming the data2 is in the next line - adjust as needed
      return lines[i + 1]?.trim() || "";
    }
  }
  return null;
}

app.post(
  "/upload",
  upload.fields([
    { name: "pdf", maxCount: 1 },
    { name: "excel", maxCount: 1 },
  ]),
  async (req, res) => {
    try {
      if (!req.files?.pdf || !req.files?.excel) {
        const error = new Error("Both PDF and Excel files are required");
        error.statusCode = 400;
        throw error;
      }

      // 1. Extract text from PDF
      const pdfText = await extractTextFromPDF(req.files.pdf[0].buffer);

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
          const foundData = findDataInPDF(pdfText, searchTerm);
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
        success: false
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