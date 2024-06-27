const express = require("express");
const XLSX = require("xlsx");
const axios = require("axios");
const multer = require("multer");
const fs = require("fs");
const path = require("path");

const app = express();
const PORT = process.env.PORT || 3000;
const upload = multer({ dest: "uploads/" });

app.use(express.static("public"));

app.get("/", (req, res) => {
  res.send(`
    <h2 style="align-item: center">Fetch API </h2>
    <button onclick="fetchData()">Fetch Data</button>
    <div id="dataDisplay"></div>
    
    <h2>Upload Excel File</h2>
    <form action="/upload" method="post" enctype="multipart/form-data">
      <input type="file" name="excelFile" accept=".xlsx,.xls">
      <input type="submit" value="Upload">
    </form>
    
    <script>
      function fetchData() {
        fetch('/fetchData')
          .then(response => response.text())
          .then(html => {
            document.getElementById('dataDisplay').innerHTML = html;
          });
      }
    </script>
  `);
});

app.get("/fetchData", (req, res) => {
  fetchAndDisplayData(res);
});

app.get("/downloadXLSX", (req, res) => {
  downloadXLSX(req, res);
});

app.post("/upload", upload.single("excelFile"), (req, res) => {
  if (!req.file) {
    return res.status(400).send("No file uploaded.");
  }

  const workbook = XLSX.readFile(req.file.path);
  const sheetName = workbook.SheetNames[0];
  const worksheet = workbook.Sheets[sheetName];
  const jsonData = XLSX.utils.sheet_to_json(worksheet);

  fs.unlinkSync(req.file.path);

  const tableHtml = convertToTable(jsonData);
  const isPostman = req.get("User-Agent").includes("PostmanRuntime");

  if (isPostman) {
    res.json(jsonData);
  } else {
    res.send(`
        ${tableHtml}
        <br>
        <a href="/downloadXLSX">Download Uploaded XLSX</a>
      `);
  }
});

async function fetchAndDisplayData(res) {
  try {
    const response = await axios.get(
      "https://mocki.io/v1/a18684c2-eb40-422a-b74d-d5a70e9a7931"
    );
    const jsonData = response.data;
    const tableHtml = convertToTable(jsonData);
    res.send(tableHtml);
  } catch (error) {
    console.error("Error in fetching data:", error);
    res.status(500).send("Error in fetching data");
  }
}

function convertToTable(jsonData) {
  if (!Array.isArray(jsonData) || jsonData.length === 0) {
    return "No data";
  }

  const headers = Object.keys(jsonData[0]);
  let tableHtml = '<table border="1"><tr>';

  headers.forEach((header) => {
    tableHtml += `<th>${header}</th>`;
  });
  tableHtml += "</tr>";

  jsonData.forEach((row) => {
    tableHtml += "<tr>";
    headers.forEach((header) => {
      tableHtml += `<td>${row[header] || ""}</td>`;
    });
    tableHtml += "</tr>";
  });

  tableHtml += "</table>";
  tableHtml += '<br><a href="/downloadXLSX">Download XLSX</a>';

  return tableHtml;
}

async function downloadXLSX(req, res) {
  try {
    const response = await axios.get(
      "https://mocki.io/v1/a18684c2-eb40-422a-b74d-d5a70e9a7931"
    );
    const jsonData = response.data;
    const workbook = XLSX.utils.book_new();
    const worksheet = XLSX.utils.json_to_sheet(jsonData);
    XLSX.utils.book_append_sheet(workbook, worksheet, "Data");
    const isPostman =
      req.headers["user-agent"] &&
      req.headers["user-agent"].includes("PostmanRuntime");

    if (isPostman) {
      const csvData = XLSX.utils.sheet_to_csv(worksheet);
      res.setHeader("Content-Type", "text/csv");
      res.setHeader("Content-Disposition", "attachment; filename=data.csv");
      return res.send(csvData);
    } else {
      const xlsxBuffer = XLSX.write(workbook, {
        type: "buffer",
        bookType: "xlsx",
      });
      res.setHeader("Content-Disposition", "attachment; filename=data.xlsx");
      res.setHeader(
        "Content-Type",
        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
      );
      return res.send(xlsxBuffer);
    }
  } catch (error) {
    console.error("Error generating file:", error);
    res.status(500).send("Error generating file");
  }
}

app.listen(PORT, () => {
  console.log(`Server running at http://localhost:${PORT}`);
});
