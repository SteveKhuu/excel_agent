const express = require('express');
const https = require('https');
const fs = require('fs');
const path = require('path');

const app = express();
const PORT = 3001;

// Middleware
app.use(express.json());
app.use((req, res, next) => {
  res.header('Access-Control-Allow-Origin', 'https://localhost:3000');
  res.header('Access-Control-Allow-Methods', 'GET, POST, PUT, DELETE, OPTIONS');
  res.header('Access-Control-Allow-Headers', 'Origin, X-Requested-With, Content-Type, Accept, Authorization');
  if (req.method === 'OPTIONS') {
    res.sendStatus(200);
  } else {
    next();
  }
});

// Proxy endpoint for Anthropic API
app.post('/api/claude', async (req, res) => {
  try {
    const { apiKey, prompt } = req.body;
    
    if (!apiKey) {
      return res.status(400).json({ error: { message: 'API key is required' } });
    }

    const response = await fetch('https://api.anthropic.com/v1/messages', {
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01'
      },
      body: JSON.stringify({
        model: 'claude-3-5-sonnet-20241022',
        max_tokens: 2000,
        messages: [
          {
            role: 'user',
            content: prompt
          }
        ]
      })
    });

    if (!response.ok) {
      const errorData = await response.json();
      return res.status(response.status).json({ error: errorData.error });
    }

    const data = await response.json();
    res.json({ content: data.content[0].text });

  } catch (error) {
    console.error('Proxy error:', error);
    res.status(500).json({ error: { message: 'Internal server error' } });
  }
});

// Get SSL certificates for HTTPS
async function getSSLOptions() {
  try {
    const devCerts = require('office-addin-dev-certs');
    const httpsOptions = await devCerts.getHttpsServerOptions();
    return {
      key: httpsOptions.key,
      cert: httpsOptions.cert
    };
  } catch (error) {
    console.error('Could not get SSL certificates:', error);
    return null;
  }
}

// Start server
async function startServer() {
  const sslOptions = await getSSLOptions();
  
  if (sslOptions) {
    // HTTPS server
    https.createServer(sslOptions, app).listen(PORT, () => {
      console.log(`Proxy server running on https://localhost:${PORT}`);
    });
  } else {
    // Fallback to HTTP
    app.listen(PORT, () => {
      console.log(`Proxy server running on http://localhost:${PORT}`);
    });
  }
}

startServer();
