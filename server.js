const express = require("express");
const cors = require("cors");
const axios = require("axios");
const fs = require("fs");
const xlsx = require("xlsx");
const app = express();
const port = 5000;

require("dotenv").config();

app.use(cors());
app.set("trust proxy", true);

app.get("/api/track", async (_, res) => {
  try {
    const ipAddress = await axios.get(process.env.IP_API_URL);
    const ip = ipAddress.data.ip;

    const { data } = await axios.get(`${process.env.GEO_API}/${ip}/json/`);
    const {
      ip: userIpAddress,
      network,
      city,
      region,
      country_name,
      latitude,
      longitude,
      org,
      country_calling_code,
    } = data;

    const visitorData = {
      ip: userIpAddress,
      network,
      city,
      region,
      country_name,
      latitude,
      longitude,
      org,
      country_calling_code,
      date: new Date().toISOString(),
    };

    // Check if the Excel file exists
    const filePath = "./visitor_data.xlsx";
    let workbook;
    if (fs.existsSync(filePath)) {
      workbook = xlsx.readFile(filePath);
    } else {
      workbook = xlsx.utils.book_new();
      const worksheet = xlsx.utils.json_to_sheet([]);
      xlsx.utils.book_append_sheet(workbook, worksheet, "Visitors");
    }

    // Get the first worksheet
    const worksheet = workbook.Sheets[workbook.SheetNames[0]];

    // Append new data to the worksheet
    const workSheetData = xlsx.utils.sheet_to_json(worksheet);
    workSheetData.push(visitorData);
    const newWorksheet = xlsx.utils.json_to_sheet(workSheetData);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    // Write to the Excel file
    xlsx.writeFile(workbook, filePath);

    res.status(204).send();
  } catch (error) {
    res.status(500).json({ error: "Something went wrong..." });
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
