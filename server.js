const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');
const { getHttpsServerOptions } = require('office-addin-dev-certs');

const app = express();
const port = 3000;

// Serve static files
app.use(express.static(__dirname));

// Get HTTPS options
const httpsOptions = getHttpsServerOptions();

// Create HTTPS server
https.createServer(httpsOptions, app).listen(port, () => {
    console.log(`Server running at https://localhost:${port}`);
}); 