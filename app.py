#!/usr/bin/env python3
"""
Vercel Entry Point for Email Classification Dashboard
"""

from web_app import app

# This is required for Vercel deployment
app = app

if __name__ == '__main__':
    app.run()