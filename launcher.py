import os
import sys
import webbrowser
import threading
import time
from app import create_app
from waitress import serve

def open_browser():
    """Open browser after a short delay"""
    time.sleep(1.5)  # Wait for server to start
    webbrowser.open('http://127.0.0.1:5000')

def run_app():
    """Run the Flask application"""
    app = create_app()
    
    # Create uploads directory if it doesn't exist
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    
    # Start browser thread
    threading.Thread(target=open_browser).start()
    
    # Run the app with waitress
    serve(app, host='127.0.0.1', port=5000)

if __name__ == '__main__':
    run_app() 