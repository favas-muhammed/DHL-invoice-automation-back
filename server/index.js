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
    origin: "http://localhost:5173",
    methods: ["GET", "POST"],
    credentials: true,
  })
);

app.use(express.json());

const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 10 * 1024 * 1024 },
});

async function extractTextFromPDF(pdfBuffer) {
  try {
    const data = await pdfParse(pdfBuffer);
    return data.text;
  } catch (error) {
    throw new Error("Failed to extract text from PDF: " + error.message);
  }
}

function formatAmount(amount) {
  const number = parseFloat(amount);
  return number.toLocaleString("en-US", {
    minimumFractionDigits: 2,
    maximumFractionDigits: 2,
  });
}

function findDataInPDF(pdfText, searchTerm) {
  const lines = pdfText.split("\n");
  for (let i = 0; i < lines.length; i++) {
    if (lines[i].includes(searchTerm)) {
      let j = i;
      let lastAmount = null;
      let previousAmount = null;
      while (j < lines.length) {
        if (
          j !== i &&
          (lines[j].match(/\d{10}/) || lines[j].includes("Total facture"))
        ) {
          break;
        }
        const amountMatch = lines[j].match(
          /(\d{1,3}(?:,\d{3})*\.\d{2}|\d+\.\d{2})\s*$/
        );
        if (amountMatch) {
          previousAmount = lastAmount;
          lastAmount = amountMatch[1].replace(/,/g, "");
        }
        if (lines[j].startsWith("Sous-Total Service") && previousAmount) {
          return formatAmount(previousAmount);
        }
        j++;
      }
      return lastAmount ? formatAmount(lastAmount) : null;
    }
  }
  return null;
}

function columnToIndex(column) {
  return (
    column
      .split("")
      .reduce(
        (acc, char) =>
          acc * 26 + char.toUpperCase().charCodeAt(0) - "A".charCodeAt(0) + 1,
        0
      ) - 1
  );
}

async function processExcelWithSeparators(
  buffer,
  referenceColumn,
  calculateTotals = false,
  totalColumn = ""
) {
  const workbook = XLSX.read(buffer);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

  const refColIndex = columnToIndex(referenceColumn);
  const totalColIndex = calculateTotals
    ? columnToIndex(totalColumn)
    : refColIndex;

  const newRows = [];
  let currentGroup = null;

  for (let i = 0; i < rows.length; i++) {
    const row = rows[i];
    const value = row[refColIndex];

    if (currentGroup !== value) {
      if (currentGroup !== null) {
        newRows.push(Array(row.length).fill(""));
        newRows.push(Array(row.length).fill(""));
      }
      currentGroup = value;
    }
    newRows.push(row);
  }

  if (calculateTotals && totalColumn) {
    let startIndex = 0;
    let currentGroup = null;

    for (let i = 0; i < newRows.length; i++) {
      const value = newRows[i][refColIndex];

      if (value === "" || i === newRows.length - 1) {
        if (startIndex < i) {
          let total = 0;
          for (let j = startIndex; j < i; j++) {
            const cellValue = newRows[j][totalColIndex];
            if (cellValue && !isNaN(parseFloat(cellValue))) {
              total += parseFloat(cellValue);
            }
          }
          if (total > 0) {
            const totalRow = Array(newRows[0].length).fill("");
            totalRow[totalColIndex] = total.toFixed(2);
            newRows.splice(i, 0, totalRow);
            i++;
          }
        }
        startIndex = i + 1;
      }
    }
  }

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
      req.body.referenceColumn,
      req.body.calculateTotals === "true",
      req.body.totalColumn
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

app.post(
  "/upload",
  upload.fields([
    { name: "pdfs", maxCount: 100 },
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

      let combinedPdfText = "";
      for (const pdfFile of req.files.pdfs) {
        const pdfText = await extractTextFromPDF(pdfFile.buffer);
        combinedPdfText += pdfText + "\n";
      }

      const workbook = XLSX.read(req.files.excel[0].buffer);
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      // Add the new column headers
      if (rows.length > 0) {
        const lastColIndex = rows[0].length;
        rows[0][lastColIndex] = "Ratio";
        rows[0][lastColIndex + 1] = "Shipping cost per item - actual amount";
        rows[0][lastColIndex + 2] = "Shipping cost - paid by client";
        rows[0][lastColIndex + 3] = "Shipping cost - gain/ loss";
      }

      // Ensure all rows have the same length
      for (let i = 1; i < rows.length; i++) {
        while (rows[i].length < rows[0].length) {
          rows[i].push(""); // Add empty cells for new columns
        }
      }

      const updates = [];

      for (let i = 1; i < rows.length; i++) {
        const searchTerm = rows[i][0];
        if (searchTerm) {
          const foundData = findDataInPDF(combinedPdfText, searchTerm);
          if (foundData) {
            rows[i][1] = foundData;
            updates.push([searchTerm, foundData]);
          }
        }
      }

      const newWorkbook = XLSX.utils.book_new();
      const newWorksheet = XLSX.utils.aoa_to_sheet(rows);
      XLSX.utils.book_append_sheet(newWorkbook, newWorksheet, sheetName);

      const excelBuffer = XLSX.write(newWorkbook, {
        type: "buffer",
        bookType: "xlsx",
      });

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

app.use(errorHandler);

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});

process.on("unhandledRejection", (err) => {
  console.error("Unhandled Promise Rejection:", err);
  process.exit(1);
});
