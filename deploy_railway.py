#!/usr/bin/env python3
"""
Simple deployment helper for Railway
"""
import os
import subprocess
import sys

def check_railway_cli():
    """Check if Railway CLI is installed"""
    try:
        result = subprocess.run(['railway', '--version'], capture_output=True, text=True)
        if result.returncode == 0:
            print("✅ Railway CLI is installed")
            return True
        else:
            print("❌ Railway CLI not found")
            return False
    except FileNotFoundError:
        print("❌ Railway CLI not installed")
        return False

def install_railway_cli():
    """Install Railway CLI"""
    print("Installing Railway CLI...")
    try:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'railway'], check=True)
        print("✅ Railway CLI installed successfully")
        return True
    except subprocess.CalledProcessError:
        print("❌ Failed to install Railway CLI")
        return False

def main():
    print("🚀 Railway Deployment Helper")
    print("=" * 40)
    
    if not check_railway_cli():
        if not install_railway_cli():
            print("\n❌ Please install Railway CLI manually:")
            print("   npm install -g @railway/cli")
            print("   or")
            print("   pip install railway")
            return
    
    print("\n📋 Next Steps:")
    print("1. Run: railway login")
    print("2. Run: railway init")
    print("3. Run: railway up")
    print("\n🌐 After deployment, update your Excel add-in with the new URL")

if __name__ == "__main__":
    main() 