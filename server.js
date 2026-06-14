/*
  Minimal HTTPS static file server for local add-in development.
  Serves this project folder at https://localhost:3000.

  It reads the trusted localhost certificate created by office-addin-dev-certs
  (run `npm run setup-certs` once before first use). PowerPoint will only load
  add-in content over HTTPS from a trusted certificate, which is why a plain
  http server will not work here.
*/
const https = require("https");
const fs = require("fs");
const path = require("path");
const os = require("os");

const PORT = 3000;
const ROOT = __dirname;
const CERT_DIR = path.join(os.homedir(), ".office-addin-dev-certs");

const TYPES = {
  ".html": "text/html; charset=utf-8",
  ".js": "text/javascript; charset=utf-8",
  ".css": "text/css; charset=utf-8",
  ".png": "image/png",
  ".xml": "application/xml",
  ".json": "application/json",
  ".svg": "image/svg+xml",
};

let options;
try {
  options = {
    key: fs.readFileSync(path.join(CERT_DIR, "localhost.key")),
    cert: fs.readFileSync(path.join(CERT_DIR, "localhost.crt")),
  };
} catch (e) {
  console.error(
    "Could not read dev certs. Run `npm run setup-certs` first.\n" + e.message
  );
  process.exit(1);
}

https
  .createServer(options, (req, res) => {
    let urlPath = decodeURIComponent(req.url.split("?")[0]);
    if (urlPath === "/") urlPath = "/src/index.html";

    // Prevent path traversal.
    const filePath = path.normalize(path.join(ROOT, urlPath));
    if (!filePath.startsWith(ROOT)) {
      res.writeHead(403);
      res.end("Forbidden");
      return;
    }

    fs.readFile(filePath, (err, data) => {
      if (err) {
        res.writeHead(404);
        res.end("Not found");
        return;
      }
      const ext = path.extname(filePath).toLowerCase();
      res.writeHead(200, {
        "Content-Type": TYPES[ext] || "application/octet-stream",
        // Allow the add-in surface itself to be framed by the Office host.
        "Content-Security-Policy": "frame-ancestors 'self' https://*.officeapps.live.com https://*.microsoft.com",
        // Stop Office from caching the add-in content so edits show on reopen
        // without a manual cache clear.
        "Cache-Control": "no-store, no-cache, must-revalidate",
        "Pragma": "no-cache",
        "Expires": "0",
      });
      res.end(data);
    });
  })
  .listen(PORT, () => {
    console.log(`Web Slide add-in served at https://localhost:${PORT}`);
    console.log("Leave this running while you use the add-in in PowerPoint.");
  });
