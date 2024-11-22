from flask import Flask
from app.config import Config
import os
import shutil

def create_app():
    app = Flask(__name__)
    app.config.from_object(Config)
    
    # Clean and recreate upload directory
    if os.path.exists(app.config['UPLOAD_FOLDER']):
        shutil.rmtree(app.config['UPLOAD_FOLDER'])
    os.makedirs(app.config['UPLOAD_FOLDER'])
    
    from app.routes import main
    app.register_blueprint(main)
    
    return app 