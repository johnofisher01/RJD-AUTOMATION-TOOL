// Simple Express server starter with CORS

const express = require('express');
const cors = require('cors');
require('dotenv').config();

const app = express();
app.use(cors());
app.use(express.json());

// Example test endpoint
app.get('/api/hello', (req, res) => {
  res.json({ message: 'Server is up and running!' });
});

// Health check
app.get('/api/health', (req, res) => {
  res.status(200).json({ status: 'ok' });
});

const PORT = process.env.PORT || 5000;
app.listen(PORT, () => {
  console.log(`Backend listening at http://localhost:${PORT}`);
});