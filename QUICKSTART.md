# Quick Start Guide

## 🚀 Get Running in 5 Minutes

### 1. Initial Setup (One-time)
```bash
# Run the setup script
./setup.sh

# Update .env with your Gemini AI API key
echo "GEMINI_API_KEY=your_actual_api_key" > .env
```

### 2. Start Everything
```bash
# Terminal 1: Start Backend
python3 main.py

# Terminal 2: Start LocalTunnel
npx localtunnel --port 8000

# Terminal 3: Start Frontend
cd "My Office Add-in"
npm start
```

### 3. Use in Excel
1. Open Excel
2. Go to Insert → Add-ins → My Add-ins → Shared Folder
3. Load your add-in
4. Click "Insert Sample Data"
5. Click "Score Leads"

## 🔧 Quick Commands

```bash
# Test if everything is working
curl https://your-tunnel-url.loca.lt/health

# Rebuild frontend after changes
cd "My Office Add-in" && npm run build

# Kill processes if needed
pkill -f "python3 main.py"
pkill -f localtunnel
```

## 🆘 Common Issues

**"Load failed"**: Insert sample data first, then score leads
**Network errors**: Use "Test Network" button in the add-in
**Tunnel expired**: Restart LocalTunnel and update URLs in code

## 📖 Full Documentation
See `README.md` for detailed setup and troubleshooting. 