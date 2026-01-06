const express = require('express');
const path = require('path');

const app = express();
const PORT = 5000;

// Disable caching for development
app.use((req, res, next) => {
  res.set('Cache-Control', 'no-cache, no-store, must-revalidate');
  res.set('Pragma', 'no-cache');
  res.set('Expires', '0');
  next();
});

// Serve static files from src directory
app.use(express.static(path.join(__dirname, 'src')));

// Serve manifest.xml at root
app.get('/manifest.xml', (req, res) => {
  res.sendFile(path.join(__dirname, 'manifest.xml'));
});

// Default route serves the taskpane
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'src', 'taskpane', 'taskpane.html'));
});

app.get('/taskpane.html', (req, res) => {
  res.sendFile(path.join(__dirname, 'src', 'taskpane', 'taskpane.html'));
});

app.listen(PORT, '0.0.0.0', () => {
  console.log(`Word Synonym Add-in server running at http://0.0.0.0:${PORT}`);
  console.log(`Manifest available at http://0.0.0.0:${PORT}/manifest.xml`);
});
