// src/server.js
import express from "express";
import path from "path";
import { fileURLToPath } from "url";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const app = express();
const PORT = 3000;

// If server.js is inside /src, don't append another /src
const SRC_DIR =
  path.basename(__dirname) === "src"
    ? __dirname
    : path.join(__dirname, "src");

// Serve static files from /src
app.use(express.static(SRC_DIR));

// Home route -> coverage chart demo (served from /src/Modules/...)
app.get("/", (req, res) => {
  res.sendFile(path.join(SRC_DIR, "Modules", "CoverageChart", "CoverageChartDemo.html"));
});

app.listen(PORT, () => {
  console.log(`Server running at http://127.0.0.1:${PORT}`);
});
