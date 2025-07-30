# Excel Claude Assistant

An AI-powered Excel add-in that uses Claude AI to analyze data, create formulas, and generate insights directly in your spreadsheets.

## 🚀 Features

- **Data Analysis**: Analyze selected data ranges and get AI-powered insights
- **Formula Creation**: Generate Excel formulas based on natural language descriptions
- **Data Insights**: Get business insights and calculated columns
- **Custom Requests**: Send custom prompts to Claude AI
- **Auto-categorization**: Automatically categorize and clean data (coming soon)
- **Quick Analysis**: Ribbon button for instant data analysis

## 📋 Prerequisites

Before setting up this project, make sure you have:

- **Node.js** (version 14 or later)
- **Microsoft Excel** (Office 365 or Excel 2016+)
- **Anthropic Claude API Key** ([Get one here](https://console.anthropic.com/))
- **Git** for cloning the repository

## 🛠️ Setup Instructions

### 1. Clone and Install

```bash
# Clone the repository
git clone <your-repository-url>
cd tryshortkhuut

# Install dependencies
npm install
```

### 2. Start the Proxy Server

This add-in requires a proxy server to handle Claude API calls due to CORS restrictions in Office Add-ins.

```bash
# Create a simple proxy server (server.js)
cat > server.js << 'EOF'
const express = require('express');
const https = require('https');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

app.post('/api/claude', async (req, res) => {
  try {
    const { apiKey, prompt } = req.body;
    
    const data = JSON.stringify({
      model: "claude-3-sonnet-20240229",
      max_tokens: 1000,
      messages: [
        {
          role: "user",
          content: prompt
        }
      ]
    });

    const options = {
      hostname: 'api.anthropic.com',
      port: 443,
      path: '/v1/messages',
      method: 'POST',
      headers: {
        'Content-Type': 'application/json',
        'x-api-key': apiKey,
        'anthropic-version': '2023-06-01',
        'Content-Length': data.length
      }
    };

    const claudeReq = https.request(options, (claudeRes) => {
      let responseData = '';
      
      claudeRes.on('data', (chunk) => {
        responseData += chunk;
      });
      
      claudeRes.on('end', () => {
        try {
          const parsed = JSON.parse(responseData);
          if (parsed.content && parsed.content[0]) {
            res.json({ content: parsed.content[0].text });
          } else {
            res.status(500).json({ error: { message: 'Unexpected response format' } });
          }
        } catch (error) {
          res.status(500).json({ error: { message: 'Failed to parse response' } });
        }
      });
    });

    claudeReq.on('error', (error) => {
      res.status(500).json({ error: { message: error.message } });
    });

    claudeReq.write(data);
    claudeReq.end();
    
  } catch (error) {
    res.status(500).json({ error: { message: error.message } });
  }
});

const PORT = process.env.PORT || 3001;
app.listen(PORT, () => {
  console.log(`Proxy server running on port ${PORT}`);
});
EOF

# Install additional dependency for proxy
npm install cors

# Start the proxy server (keep this running)
node server.js
```

### 3. Build and Start the Add-in

Open a new terminal window and run:

```bash
# Build the project
npm run build:dev

# Start the development server
npm run dev-server
```

The add-in will be served at `https://localhost:3000`

### 4. Install SSL Certificates

Office Add-ins require HTTPS. Generate development certificates:

```bash
# Generate certificates (this should happen automatically)
npx office-addin-dev-certs install

# If you encounter certificate issues, try:
npx office-addin-dev-certs install --machine
```

### 5. Sideload the Add-in

#### Option A: Automatic Sideloading (Recommended)
```bash
# Start debugging (this will open Excel and sideload the add-in)
npm start
```

#### Option B: Manual Sideloading
1. Open Excel
2. Go to **Insert** → **Add-ins** → **My Add-ins**
3. Click **Upload My Add-in**
4. Select the `manifest.xml` file from your project root
5. Click **Upload**

### 6. Configure Your API Key

1. In Excel, look for the **Claude AI** tab in the ribbon
2. Click **Show Taskpane** to open the add-in panel
3. Enter your Anthropic Claude API key in the API Key field
4. The key will be saved automatically for future use

## 🎯 How to Use

### Quick Analysis (Ribbon Button)
1. Select a data range in Excel
2. Click the **Analyze Data** button in the Claude AI ribbon
3. A new worksheet will be created with quick analysis results

### Detailed Analysis (Task Pane)
1. Open the Claude AI task pane from the ribbon
2. Select your data range
3. Choose from available actions:
   - **Analyze Selection**: Get detailed insights and suggested new columns
   - **Create Formula**: Generate formulas based on your description
   - **Data Insights**: Get business insights with calculated columns
   - **Custom Request**: Send any custom prompt to Claude

### Example Workflows

#### 1. Sales Data Analysis
```
1. Select your sales data range
2. Click "Analyze Selection"
3. Claude will suggest new columns like:
   - Revenue per customer
   - Month-over-month growth
   - Sales performance categories
```

#### 2. Formula Creation
```
1. Select your data
2. Click "Create Formula"
3. Describe what you need: "Calculate the percentage change from last month"
4. Claude generates the Excel formula and applies it
```

#### 3. Custom Analysis
```
1. Enter a custom prompt like: "Create a financial projection table for the next 5 years"
2. Click "Send Custom Request"
3. Claude creates a new worksheet with the requested analysis
```

## 📁 Project Structure

```
tryshortkhuut/
├── src/
│   ├── commands/
│   │   ├── commands.html          # Ribbon command functions
│   │   └── commands.js
│   └── taskpane/
│       ├── taskpane.html          # Main UI
│       ├── taskpane.js            # Main functionality
│       └── taskpane.css           # Styling
├── assets/                        # Icons and images
├── manifest.xml                   # Add-in manifest
├── webpack.config.js              # Build configuration
├── package.json                   # Dependencies and scripts
├── server.js                      # Proxy server
└── README.md
```

## 🔧 Development Commands

```bash
# Build for development
npm run build:dev

# Build for production
npm run build

# Start development server
npm run dev-server

# Watch for changes
npm run watch

# Lint code
npm run lint

# Fix linting issues
npm run lint:fix

# Start debugging in Excel
npm start

# Stop debugging
npm stop

# Validate manifest
npm run validate
```

## 🚨 Troubleshooting

### Common Issues

#### 1. Certificate Errors
```bash
# Reinstall certificates
npx office-addin-dev-certs uninstall
npx office-addin-dev-certs install

# For Mac users, you might need:
sudo npx office-addin-dev-certs install --machine
```

#### 2. Add-in Not Loading
- Ensure both the dev server (`npm run dev-server`) and proxy server (`node server.js`) are running
- Check that ports 3000 and 3001 are not blocked by firewall
- Try clearing Excel's cache: Close Excel, delete `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\` (Windows) or `~/Library/Containers/com.microsoft.Excel/Data/Documents/wef/` (Mac)

#### 3. API Errors
- Verify your Claude API key is correct and has credits
- Check the proxy server console for error messages
- Ensure you have internet connectivity

#### 4. CORS Issues
- Make sure the proxy server is running on port 3001
- The proxy server handles CORS for Claude API calls
- Don't try to call Claude API directly from the add-in

### Development Tips

1. **Hot Reload**: Use `npm run watch` for automatic rebuilding during development
2. **Debugging**: Use browser dev tools in the task pane (right-click → Inspect)
3. **Console Logs**: Check both Excel's console and the proxy server console for errors
4. **Testing**: Test with small data ranges first, then scale up

## 🌐 API Configuration

The add-in uses Claude AI through a proxy server to avoid CORS issues. The proxy server:
- Runs on `localhost:3001`
- Forwards requests to `https://api.anthropic.com`  
- Handles authentication and response formatting
- Supports Claude 3 Sonnet model

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)  
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📚 Additional Resources

- [Office Add-ins Documentation](https://docs.microsoft.com/en-us/office/dev/add-ins/)
- [Claude AI API Documentation](https://docs.anthropic.com/claude/reference/getting-started-with-the-api)
- [Excel JavaScript API Reference](https://docs.microsoft.com/en-us/javascript/api/excel)

## 🔄 Updates and Maintenance

To update the add-in:
1. Pull the latest changes from the repository
2. Run `npm install` to update dependencies  
3. Rebuild with `npm run build:dev`
4. Restart the development server

For production deployment, build with `npm run build` and deploy the `dist/` folder to your web server.
