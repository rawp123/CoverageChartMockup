const express = require("express");
const path = require("path");

const app = express();
const PORT = 3000;

// Serve everything inside /src as web root
app.use(express.static(path.join(__dirname, "src")));

// Load demo at root URL
app.get("/", (req, res) => {
  res.sendFile(
    path.join(__dirname, "src", "Modules", "CoverageChart", "CoverageChartDemo.html")
  );
});

app.listen(PORT, () => {
  console.log(`Server running at http://127.0.0.1:${PORT}`);
});
