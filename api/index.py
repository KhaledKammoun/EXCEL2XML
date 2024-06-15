from flask import Flask, request, jsonify
import io
import pandas as pd
import xml.etree.ElementTree as ET
import sqlite3

app = Flask(__name__)
@app.route('/')
def home() :
    return "HELLO, THIS IS EXCEL2XML"

@app.route('/about')
def about() :
    return "Hello, /about test routing"

def xml_to_excel(xml_content, excel_path):
    tree = ET.ElementTree(ET.fromstring(xml_content))
    root = tree.getroot()
    
    data = []
    for child in root:
        data.append(child.attrib)

    df = pd.DataFrame(data)
    df.to_excel(excel_path, index=False)

def excel_to_xml(excel_content, xml_path):
    df = pd.read_excel(io.BytesIO(excel_content))
    root = ET.Element("root")
    
    for _, row in df.iterrows():
        elem = ET.SubElement(root, "item", attrib=row.to_dict())

    tree = ET.ElementTree(root)
    tree.write(xml_path)

def store_file_in_db(file_content, file_type):
    conn = sqlite3.connect('files.db')
    cursor = conn.cursor()

    cursor.execute('''CREATE TABLE IF NOT EXISTS files
                      (id INTEGER PRIMARY KEY, file_type TEXT, file BLOB)''')

    cursor.execute("INSERT INTO files (file_type, file) VALUES (?, ?)", (file_type, file_content))

    conn.commit()
    conn.close()

@app.route('/convert', methods=['POST'])
def convert():
    file = request.files['file']
    file_type = request.form['file_type']
    output_type = request.form['output_type']
    
    file_content = file.read()
    
    if file_type == 'xml' and output_type == 'excel':
        xml_to_excel(file_content, 'output.xlsx')
    elif file_type == 'excel' and output_type == 'xml':
        excel_to_xml(file_content, 'output.xml')
    
    store_file_in_db(file_content, file_type)
    
    return jsonify({'status': 'success'})

if __name__ == '__main__':
    app.run()
