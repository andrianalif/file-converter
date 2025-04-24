from flask import Flask, request, jsonify
from flask_cors import CORS
import os
import base64
import requests
from openpyxl import load_workbook
from docx import Document
from PyPDF2 import PdfReader
import io
from dotenv import load_dotenv

app = Flask(__name__)
CORS(app)

# Load environment variables
load_dotenv()

# WordPress credentials
WP_USER = "alif.adrian"
WP_APP_PASSWORD = "xBa8 zJjb UDhK oXcP NeqR 7MDm"
WP_API_URL = "https://frontier.vstecs.ai/wp-json/wp/v2/pages"

def convert_excel_to_html(file_content):
    wb = load_workbook(io.BytesIO(file_content), data_only=False)
    html_content = """
    <style>
        body { 
            font-family: 'Segoe UI', sans-serif; 
            line-height: 1.6; 
            color: #333; 
            margin: 0; 
            padding: 0; 
            background-color: #f4f4f4; 
        }
        h1, h2 { 
            color: #2c3e50; 
            margin-bottom: 20px;
        }
        table { 
            width: 100%; 
            border-collapse: collapse; 
            margin-bottom: 40px; 
            background-color: #fff; 
            border: 1px solid #ddd;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        th, td { 
            border: 1px solid #ccc; 
            padding: 12px; 
            text-align: left; 
            vertical-align: top; 
        }
        th { 
            background-color: #f4f4f4; 
            font-weight: bold;
            color: #2c3e50;
        }
        tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        tr:hover {
            background-color: #f1f1f1;
        }
        hr { 
            border: 0; 
            height: 1px; 
            background: #ddd; 
            margin: 40px 0; 
        }
        p { 
            font-size: 16px;
            margin-bottom: 15px;
        }
        code { 
            font-size: 14px; 
            background-color: #f9f9f9; 
            padding: 2px 6px; 
            border-radius: 4px;
            font-family: monospace;
        }
        .table-container { 
            padding: 20px;
            background-color: white;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .table-wrapper { 
            margin-bottom: 40px;
            overflow-x: auto;
        }
        .price-value {
            text-align: right;
            font-family: monospace;
        }
        .currency {
            color: #2c3e50;
            font-weight: bold;
        }
    </style>
    
    <h1>Arista Networks Official Product Price List</h1>
    <p>This comprehensive listing includes various Arista product lines such as switches, software licenses, cloud subscriptions, and accessories. Each table below represents a different product category from the corresponding Excel sheets.</p>
    <p>All prices are displayed in <strong>US Dollars (USD)</strong>, and formatting has been preserved to retain accuracy, including symbols like <code>$</code> and number grouping such as <code>,</code> or <code>.</code>.</p>
    <p>Use this information to identify and reference Arista part numbers for procurement, pricing comparison, or documentation purposes.</p>
    <hr>
    """

    for sheet in wb.sheetnames:
        ws = wb[sheet]
        html_content += f"<h2>Product Category: {sheet}</h2>\n"
        html_content += "<div class='table-container'><div class='table-wrapper'>\n"
        html_content += "<table>\n"

        for row_idx, row in enumerate(ws.iter_rows(values_only=False)):
            html_content += "<tr>"
            for cell in row:
                value = cell.value
                number_format = cell.number_format

                if value is None:
                    display_value = ""
                elif isinstance(value, (int, float)):
                    if '$' in number_format:
                        display_value = f"<span class='currency'>$</span><span class='price-value'>{value:,.2f}</span>"
                    elif ',' in number_format or '.' in number_format:
                        display_value = f"<span class='price-value'>{value:,.0f}</span>"
                    else:
                        display_value = str(value)
                else:
                    display_value = str(value)

                tag = "th" if row_idx == 0 else "td"
                html_content += f"<{tag}>{display_value}</{tag}>"
            html_content += "</tr>\n"

        html_content += "</table></div></div>\n<hr>\n"

    return html_content

def convert_docx_to_html(file_content):
    doc = Document(io.BytesIO(file_content))
    html_content = "<div class='docx-content'>"
    
    for para in doc.paragraphs:
        html_content += f"<p>{para.text}</p>"
    
    html_content += "</div>"
    return html_content

def convert_pdf_to_html(file_content):
    pdf = PdfReader(io.BytesIO(file_content))
    html_content = "<div class='pdf-content'>"
    
    for page in pdf.pages:
        html_content += f"<div class='page'>{page.extract_text()}</div>"
    
    html_content += "</div>"
    return html_content

def publish_to_wordpress(title, content):
    print("🚀 Publishing to WordPress as a PAGE...")

    credentials = f"{WP_USER}:{WP_APP_PASSWORD}"
    token = base64.b64encode(credentials.encode())
    headers = {
        "Authorization": f"Basic {token.decode('utf-8')}",
        "Content-Type": "application/json"
    }

    # First, get the parent page ID for "Price List"
    parent_response = requests.get(
        f"{WP_API_URL}?search=Price List",
        headers=headers
    )
    
    parent_id = None
    if parent_response.status_code == 200:
        pages = parent_response.json()
        for page in pages:
            if page['title']['rendered'] == "Price List":
                parent_id = page['id']
                break
    
    if not parent_id:
        print("⚠️ Warning: Parent page 'Price List' not found. Creating as a top-level page.")
    
    page = {
        "title": title,
        "content": content,
        "status": "publish",
        "parent": parent_id if parent_id else 0
    }

    response = requests.post(WP_API_URL, headers=headers, json=page)

    if response.status_code == 201:
        print(f"✅ Page published successfully! URL: {response.json().get('link')}")
        return response
    else:
        print(f"❌ Failed to publish. Status: {response.status_code}")
        print(response.text)
        return response

@app.route('/api/convert', methods=['POST'])
def convert_file():
    if 'file' not in request.files:
        return jsonify({'error': 'No file provided'}), 400
    
    file = request.files['file']
    file_type = file.content_type
    file_content = file.read()
    
    try:
        if file_type == 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet':
            html_content = convert_excel_to_html(file_content)
        elif file_type == 'application/vnd.openxmlformats-officedocument.wordprocessingml.document':
            html_content = convert_docx_to_html(file_content)
        elif file_type == 'application/pdf':
            html_content = convert_pdf_to_html(file_content)
        else:
            return jsonify({'error': 'Unsupported file type'}), 400
        
        return jsonify({'html': html_content})
    
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/api/publish', methods=['POST'])
def publish():
    data = request.json
    title = data.get('title')
    content = data.get('content')
    
    if not title or not content:
        return jsonify({'error': 'Title and content are required'}), 400
    
    response = publish_to_wordpress(title, content)
    
    if response.status_code == 201:
        return jsonify({'url': response.json().get('link')})
    else:
        return jsonify({'error': 'Failed to publish'}), response.status_code

if __name__ == '__main__':
    app.run(debug=True) 