from flask import Flask, render_template, request, redirect, send_file, url_for, flash
import os
from scripts import createXMLFile, convert_xml_to_excel

app = Flask(__name__)
app.secret_key = 'your_secret_key'

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/convert_excel_to_xml', methods=['GET', 'POST'])
def convert_excel_to_xml():
    if request.method == 'POST':
        # Handle form submission
        input_file = request.files['input_file']
        output_folder = request.form['output_folder']
        output_file_name = request.form['output_file_name']

        # Save the uploaded file
        filename = input_file.filename
        file_path = os.path.join(output_folder, filename)
        input_file.save(file_path)

        # Process conversion
        try:
            createXMLFile(file_path, os.path.join(output_folder, output_file_name))
            flash('Conversion successful!', 'success')
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')

        return redirect(url_for('index'))

    return render_template('convert_excel_to_xml.html')
from werkzeug.utils import secure_filename

@app.route('/convert_xml_to_excel', methods=['GET', 'POST'])
def convert_xml_to_excel():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'input_file' not in request.files:
            flash('No file part', 'error')
            return redirect(request.url)
        
        input_file = request.files['input_file']
        output_folder = request.form['output_folder']
        output_file_name = request.form['output_file_name']

        # Check if the file is empty
        if input_file.filename == '':
            flash('No selected file', 'error')
            return redirect(request.url)

        # Save the uploaded file
        filename = secure_filename(input_file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        input_file.save(file_path)

        # Process conversion
        try:
            convert_xml_to_excel(file_path, os.path.join(output_folder, output_file_name))
            flash('Conversion successful!', 'success')
        except Exception as e:
            flash(f'Error: {str(e)}', 'error')
        
        # Optionally, remove the uploaded file after processing
        os.remove(file_path)

        return redirect(url_for('index'))

    return render_template('convert_xml_to_excel.html')
@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    return send_file(filename, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
