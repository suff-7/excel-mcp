#!/usr/bin/env python3
"""
Test script for Excel FastMCP Server
"""

import requests
import json
import sys

def test_server(base_url):
    """Test basic server functionality"""
    print(f"Testing server at: {base_url}")
    
    try:
        # Test if server is running
        response = requests.get(f"{base_url}/", timeout=10)
        print(f"Server status: {response.status_code}")
        
        if response.status_code == 200:
            print("✓ Server is running successfully!")
            return True
        else:
            print(f"✗ Server returned status {response.status_code}")
            return False
            
    except requests.exceptions.ConnectionError:
        print("✗ Could not connect to server")
        return False
    except requests.exceptions.Timeout:
        print("✗ Server request timed out")
        return False
    except Exception as e:
        print(f"✗ Error testing server: {e}")
        return False

if __name__ == "__main__":
    if len(sys.argv) > 1:
        url = sys.argv[1]
    else:
        url = "http://localhost:8000"
    
    success = test_server(url)
    sys.exit(0 if success else 1)
