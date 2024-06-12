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

app.get("/", (_, res) => {
  res.send("Hello Welcome to my backend!");
});

app.get("/api/track", async (_, res) => {
  try {
    const ipAddress = await axios.get(process.env.IP_API_URL);
    const ip = ipAddress.data.ip;

    const { data } = await axios.get(`${process.env.GEO_API}/${ip}/json/`);
    const {
      ip: userIpAddress,
      city,
      region,
      country_name,
      latitude,
      longitude,
    } = data;

    // Check if the Excel file exists
    const currentDate = new Date().toISOString();
    const filePath = "./visitor_data.xlsx";

    // Initialize workbook and worksheet
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
    const dataVisitors = xlsx.utils.sheet_to_json(worksheet);

    // Check if the visitor's IP already exists
    const existingVisitor = dataVisitors.find(
      (visitor) => visitor.ip === userIpAddress
    );
    if (existingVisitor) {
      // Update the existing visitor's countVisited and date
      existingVisitor.countVisited += 1;
      existingVisitor.date = currentDate;
    } else {
      // Add a new visitor entry
      dataVisitors.push({
        userIpAddress,
        city,
        region,
        latitude,
        longitude,
        country: country_name,
        date: currentDate,
        countVisited: 1,
      });
    }

    // Convert JSON data back to worksheet
    const newWorksheet = xlsx.utils.json_to_sheet(dataVisitors);
    workbook.Sheets[workbook.SheetNames[0]] = newWorksheet;

    // Write to the Excel file
    xlsx.writeFile(workbook, filePath);

    res.status(204).send(); // Send no content response
  } catch (error) {
    res.status(500).json({ error: "Something went wrong..." });
  }
});

app.listen(port, () => {
  console.log(`Server is running on http://localhost:${port}`);
});
