#!/bin/bash

echo "🚀 Setting up AI-Powered Lead Scoring Excel Add-in"
echo "=================================================="

# Check if .env file exists
if [ ! -f ".env" ]; then
    echo "📝 Creating .env file..."
    echo "GEMINI_API_KEY=your_gemini_api_key_here" > .env
    echo "⚠️  Please update .env with your actual Gemini AI API key"
else
    echo "✅ .env file already exists"
fi

# Install Python dependencies
echo "🐍 Installing Python dependencies..."
pip3 install -r requirements.txt

# Install Node.js dependencies
echo "📦 Installing Node.js dependencies..."
cd "My Office Add-in"
npm install
cd ..

echo ""
echo "✅ Setup complete!"
echo ""
echo "Next steps:"
echo "1. Update .env with your Gemini AI API key"
echo "2. Start the backend: python3 main.py"
echo "3. Start LocalTunnel: npx localtunnel --port 8000"
echo "4. Start the frontend: cd 'My Office Add-in' && npm start"
echo ""
echo "📖 See README.md for detailed instructions" 