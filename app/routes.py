from flask import Blueprint, render_template, request, jsonify, current_app, send_from_directory, redirect, url_for
import os
from werkzeug.utils import secure_filename
from app.services.word_processor import WordProcessor
import win32com.client as win32
import pythoncom
from datetime import datetime

main = Blueprint('main', __name__)

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in current_app.config['ALLOWED_EXTENSIONS']

def convert_to_pdf(word_path, pdf_path):
    pythoncom.CoInitialize()
    try:
        word = win32.Dispatch('Word.Application')
        doc = word.Documents.Open(word_path)
        doc.SaveAs(pdf_path, FileFormat=17)  # 17 represents PDF format
        doc.Close()
        word.Quit()
    finally:
        pythoncom.CoUninitialize()

def ensure_upload_folder():
    """Ensure upload folder exists"""
    os.makedirs(current_app.config['UPLOAD_FOLDER'], exist_ok=True)

@main.route('/')
def index():
    ensure_upload_folder()
    return render_template('index.html')

@main.route('/process', methods=['POST'])
def process_document():
    ensure_upload_folder()
    
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    hidden_text = request.form.get('hidden_text', '')
    
    if file.filename == '':
        return jsonify({'error': 'No file selected'}), 400
    
    if file and allowed_file(file.filename):
        try:
            # Create filenames with _modified suffix
            base_filename = os.path.splitext(secure_filename(file.filename))[0]
            ext = os.path.splitext(file.filename)[1]
            
            input_filename = f"{base_filename}{ext}"
            output_filename = f"{base_filename}_modified{ext}"
            pdf_filename = f"{base_filename}_modified.pdf"
            
            input_path = os.path.join(current_app.config['UPLOAD_FOLDER'], input_filename)
            output_path = os.path.join(current_app.config['UPLOAD_FOLDER'], output_filename)
            pdf_path = os.path.join(current_app.config['UPLOAD_FOLDER'], pdf_filename)
            
            # Save uploaded file
            file.save(input_path)
            
            # Process document
            with WordProcessor() as processor:
                result = processor.process_document(input_path, output_path, hidden_text)
            
            # Convert to PDF if Word processing was successful
            if result:
                try:
                    convert_to_pdf(output_path, pdf_path)
                except Exception as e:
                    print(f"PDF conversion failed: {str(e)}")
                    # Continue even if PDF conversion fails
            
            # Clean up input file
            try:
                if os.path.exists(input_path):
                    os.remove(input_path)
            except Exception as e:
                print(f"Cleanup failed: {str(e)}")
            
            if result:
                return redirect(url_for('main.download_page', 
                                      word_file=output_filename,
                                      pdf_file=pdf_filename))
            
            return jsonify({'error': 'Processing failed'}), 500
            
        except Exception as e:
            print(f"Processing error: {str(e)}")
            return jsonify({'error': str(e)}), 500
    
    return jsonify({'error': 'Invalid file type'}), 400

@main.route('/download-page')
def download_page():
    word_file = request.args.get('word_file')
    pdf_file = request.args.get('pdf_file')
    
    # Verify files exist
    word_path = os.path.join(current_app.config['UPLOAD_FOLDER'], word_file)
    pdf_path = os.path.join(current_app.config['UPLOAD_FOLDER'], pdf_file)
    
    if not os.path.exists(word_path):
        return jsonify({'error': 'Word file not found'}), 404
        
    return render_template('download.html',
                         word_download_url=url_for('main.download_file', filename=word_file),
                         pdf_download_url=url_for('main.download_file', filename=pdf_file))

@main.route('/download/<filename>')
def download_file(filename):
    try:
        return send_from_directory(
            current_app.config['UPLOAD_FOLDER'],
            filename,
            as_attachment=True
        )
    except Exception as e:
        print(f"Download error: {str(e)}")
        return jsonify({'error': str(e)}), 404 