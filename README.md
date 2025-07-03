# AI-Powered Lead Scoring Excel Add-in

A full-stack Microsoft Excel add-in that provides AI-powered lead scoring and analysis using React, Python FastAPI, and Google Gemini AI.

## 🚀 Features

- **AI-Powered Lead Scoring**: Real-time lead quality analysis using Gemini AI
- **Smart Data Insertion**: Adds realistic sample lead data for testing
- **Visual Feedback**: Color-coded highlighting based on lead scores
- **Detailed Insights**: Individual lead analysis with recommendations and risk factors
- **Summary Dashboard**: Overview of lead distribution and quality metrics
- **Network Testing**: Built-in connectivity testing for troubleshooting

## 🏗️ Architecture

- **Frontend**: React + Office.js (Excel Add-in)
- **Backend**: Python FastAPI with Gemini AI integration
- **Tunneling**: LocalTunnel for HTTPS access from Excel
- **AI**: Google Gemini AI for intelligent lead analysis

## 📋 Prerequisites

- **Node.js** (v14 or higher)
- **Python 3.9+**
- **Microsoft Excel** (Desktop version)
- **Google Gemini AI API Key**

## 🛠️ Setup Instructions

### 1. Clone and Navigate
```bash
git clone <your-repo-url>
cd excelPlugin
```

### 2. Environment Setup
Create a `.env` file in the root directory:
```bash
# Create .env file
touch .env
```

Add your Gemini AI API key:
```env
GEMINI_API_KEY=your_gemini_api_key_here
```

### 3. Backend Setup
```bash
# Install Python dependencies
pip3 install -r requirements.txt

# Start the FastAPI backend
python3 main.py
```

The backend will start on `http://localhost:8000`

### 4. Frontend Setup
```bash
# Navigate to the add-in directory
cd "My Office Add-in"

# Install Node.js dependencies
npm install

# Build the add-in
npm run build

# Start the development server
npm start
```

The frontend will start on `https://localhost:3000`

### 5. LocalTunnel Setup
In a new terminal window:
```bash
# Start LocalTunnel for HTTPS access
npx localtunnel --port 8000
```

Note the generated URL (e.g., `https://nasty-ears-decide.loca.lt`)

### 6. Update Configuration
After getting the LocalTunnel URL, update the following files:

**Frontend API URL** (`My Office Add-in/src/taskpane/components/App.jsx`):
```javascript
// Replace the URL in the fetch calls
'https://your-new-tunnel-url.loca.lt/score-leads'
```

**Manifest** (`My Office Add-in/manifest.xml`):
```xml
<AppDomain>https://your-new-tunnel-url.loca.lt</AppDomain>
```

**Backend CORS** (`main.py`):
```python
allow_origins=[
    "https://localhost:3000",
    "https://your-new-tunnel-url.loca.lt",
    # ... other domains
]
```

### 7. Rebuild and Restart
```bash
# Rebuild frontend
cd "My Office Add-in"
npm run build
npm start

# Restart backend (if needed)
# Ctrl+C to stop, then python3 main.py
```

## 🎯 Usage

### In Excel:
1. **Load the Add-in**: Go to Insert → Add-ins → My Add-ins → Shared Folder
2. **Insert Sample Data**: Click "Insert Sample Data" to add realistic lead data
3. **Score Leads**: Click "Score Leads" to analyze with AI
4. **View Results**: Check the task pane for detailed insights
5. **Test Network**: Use "Test Network" button for connectivity testing

### Expected Workflow:
1. Sample data is inserted with headers and realistic lead information
2. AI analyzes each lead based on multiple factors (revenue, engagement, budget, etc.)
3. Scores are calculated (0-100) with color-coded highlighting
4. Detailed insights and recommendations are displayed
5. Summary statistics show lead distribution

## 🔧 Troubleshooting

### Common Issues:

**1. "Load failed" Error**
- Ensure Excel has data selected
- Try "Insert Sample Data" first
- Check console (F12) for detailed errors

**2. Network Connectivity Issues**
- Use "Test Network" button to diagnose
- Ensure LocalTunnel is running
- Check if tunnel URL has expired (restart if needed)

**3. LocalTunnel Password Page**
- The add-in includes bypass headers automatically
- If issues persist, restart LocalTunnel with a fresh URL

**4. Backend Connection Failed**
- Verify backend is running on port 8000
- Check if LocalTunnel URL is updated in code
- Ensure CORS is properly configured

### Debug Commands:
```bash
# Test backend health
curl https://your-tunnel-url.loca.lt/health

# Test scoring endpoint
curl -X POST https://your-tunnel-url.loca.lt/score-leads \
  -H "Content-Type: application/json" \
  -H "bypass-tunnel-reminder: true" \
  -d '{"leads":[{"Name":"Test","Email":"test@example.com","Company":"Test Corp"}]}'

# Check if ports are in use
lsof -i :8000
lsof -i :3000
```

## 📁 Project Structure

```
excelPlugin/
├── main.py                 # FastAPI backend with Gemini AI
├── requirements.txt        # Python dependencies
├── .env                   # Environment variables (API keys)
├── My Office Add-in/      # Excel add-in frontend
│   ├── src/
│   │   └── taskpane/
│   │       └── components/
│   │           └── App.jsx # Main React component
│   │   ├── manifest.xml       # Excel add-in manifest
│   │   ├── package.json       # Node.js dependencies
│   │   └── webpack.config.js  # Build configuration
│   └── README.md             # This file
```

## 🔒 Security Notes

- **API Keys**: Never commit `.env` files to version control
- **HTTPS**: LocalTunnel provides HTTPS for Excel compatibility
- **CORS**: Backend is configured for specific origins only
- **Network**: Excel add-ins require HTTPS for external API calls

## 🚀 Deployment

For production deployment:
1. Deploy backend to a cloud service (Railway, Render, etc.)
2. Update frontend API URLs to production endpoint
3. Update manifest with production domains
4. Package and distribute the add-in

## 📝 API Endpoints

- `GET /health` - Backend health check
- `POST /score-leads` - Lead scoring with Gemini AI
- `GET /` - Root endpoint

---

**Note**: This add-in requires a Google Gemini AI API key for full functionality. The backend includes comprehensive error handling and fallback mechanisms for testing without AI. 