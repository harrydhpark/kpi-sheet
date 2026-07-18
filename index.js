const http = require("http");
const fs = require("fs");
const path = require("path");

const PORT = 8080;

// Map file extensions to content types
const MIME_TYPES = {
    ".html": "text/html; charset=utf-8",
    ".css": "text/css; charset=utf-8",
    ".js": "application/javascript; charset=utf-8",
    ".json": "application/json; charset=utf-8",
    ".png": "image/png",
    ".jpg": "image/jpeg",
    ".gif": "image/gif",
    ".svg": "image/svg+xml",
    ".ico": "image/x-icon",
    ".txt": "text/plain; charset=utf-8"
};

const server = http.createServer((req, res) => {
    // Decode URL to handle Korean folder/file names correctly!
    let reqUrl = decodeURIComponent(req.url);
    
    // Default to index.html
    if (reqUrl === "/" || reqUrl === "") {
        reqUrl = "/index.html";
    }

    // Resolve file path
    const filePath = path.join(__dirname, reqUrl);

    // Security check: ensure the file is inside the project directory
    if (!filePath.startsWith(__dirname)) {
        res.writeHead(403, { "Content-Type": "text/plain" });
        res.end("403 Forbidden");
        return;
    }

    // Read and serve the file
    fs.stat(filePath, (err, stats) => {
        if (err || !stats.isFile()) {
            res.writeHead(404, { "Content-Type": "text/plain" });
            res.end("404 Not Found");
            return;
        }

        const ext = path.extname(filePath).toLowerCase();
        const contentType = MIME_TYPES[ext] || "application/octet-stream";

        res.writeHead(200, { "Content-Type": contentType });
        const stream = fs.createReadStream(filePath);
        stream.pipe(res);
    });
});

server.listen(PORT, "0.0.0.0", () => {
    console.log(`Static file server started on 0.0.0.0:${PORT}`);
});
