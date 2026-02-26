const express = require("express");
const path = require("path");

const app = express();

app.use("/CoverageChart", express.static(path.join(__dirname, "CoverageChart")));
app.use("/data", express.static(path.join(__dirname, "../data")));

app.get("/coverage-dashboard", (_req, res) => {
  res.sendFile(path.join(__dirname, "CoverageChartDemo.html"));
});

app.get("/", (_req, res) => res.redirect("/coverage-dashboard"));

app.listen(3000, () => console.log("Server is running on http://localhost:3000"));