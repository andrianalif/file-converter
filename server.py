from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import os
import time
import logging
import base64
from openpyxl import load_workbook
import warnings

# Suppress the WMF image warning
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# WordPress API credentials
WP_USER = "alif.adrian"
WP_APP_PASSWORD = "xBa8 zJjb UDhK oXcP NeqR 7MDm"
WP_API_URL = "https://frontier.vstecs.ai/wp-json/wp/v2/pages"

logger.info("WordPress credentials loaded successfully")

def convert_excel_to_html(file_path):
    try:
        logger.debug(f"Converting Excel file: {file_path}")
        # Load workbook with data_only=True to get calculated values instead of formulas
        wb = load_workbook(file_path, data_only=True)
        logger.debug(f"Excel file loaded successfully with sheets: {wb.sheetnames}")
        
        html_content = """
        <style>
            body { 
                font-family: 'Segoe UI', sans-serif; 
                line-height: 1.6; 
                color: #333; 
                margin: 0; 
                padding: 20px; 
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
        
        <h1>Price List</h1>
        <p>This comprehensive listing includes various product lines and their corresponding prices. Each table below represents a different product category from the corresponding Excel sheets.</p>
        <p>All prices are displayed in <strong>US Dollars (USD)</strong>, and formatting has been preserved to retain accuracy, including symbols like <code>$</code> and number grouping such as <code>,</code> or <code>.</code>.</p>
        <p>Use this information to identify and reference part numbers for procurement, pricing comparison, or documentation purposes.</p>
        <hr>
        """
        
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            html_content += f"<h2>Product Category: {sheet}</h2>\n"
            html_content += "<div class='table-container'><div class='table-wrapper'>\n"
            html_content += "<table>\n"

            # Get column widths from Excel
            col_widths = []
            for col in ws.columns:
                max_length = 0
                for cell in col:
                    if cell.value is not None:
                        max_length = max(max_length, len(str(cell.value)))
                col_widths.append(max_length)

            for row_idx, row in enumerate(ws.iter_rows(values_only=False)):
                html_content += "<tr>"
                for col_idx, cell in enumerate(row):
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
                    style = f"min-width: {col_widths[col_idx] * 8}px;"
                    html_content += f"<{tag} style='{style}'>{display_value}</{tag}>"
                html_content += "</tr>\n"

            html_content += "</table></div></div>\n<hr>\n"
        
        logger.debug("HTML conversion completed successfully")
        return html_content
    except Exception as e:
        logger.error(f"Error converting Excel to HTML: {str(e)}")
        raise Exception(f"Error converting Excel to HTML: {str(e)}")

def publish_to_wordpress(title, content):
    try:
        logger.debug(f"Publishing to WordPress with title: {title}")
        
        # First, get the parent page ID for "Price List"
        parent_url = f"{WP_API_URL}?search=Price List"
        logger.debug(f"Searching for parent page: {parent_url}")
        
        credentials = f"{WP_USER}:{WP_APP_PASSWORD}"
        token = base64.b64encode(credentials.encode())
        headers = {
            "Authorization": f"Basic {token.decode('utf-8')}",
            "Content-Type": "application/json"
        }
        
        parent_response = requests.get(
            parent_url,
            headers=headers
        )
        
        parent_id = None
        if parent_response.status_code == 200:
            pages = parent_response.json()
            for page in pages:
                if page['title']['rendered'] == "Price List":
                    parent_id = page['id']
                    logger.debug(f"Found parent page 'Price List' with ID: {parent_id}")
                    break
        
        if not parent_id:
            logger.warning("Parent page 'Price List' not found. Creating as a top-level page.")
        
        # Prepare the page data
        page_data = {
            'title': title,
            'content': content,
            'status': 'publish',
            'parent': parent_id if parent_id else 0
        }
        
        # Make the API request
        logger.debug("Sending request to WordPress API")
        response = requests.post(
            WP_API_URL,
            headers=headers,
            json=page_data
        )
        
        logger.debug(f"WordPress API response status: {response.status_code}")
        
        if response.status_code == 201:
            page_url = response.json()['link']
            logger.info(f"Page published successfully: {page_url}")
            return page_url
        else:
            error_message = response.text
            try:
                error_json = response.json()
                if 'message' in error_json:
                    error_message = error_json['message']
            except:
                pass
            logger.error(f"WordPress API error: {error_message}")
            raise Exception(f"WordPress API error: {error_message}")
    except Exception as e:
        logger.error(f"Error publishing to WordPress: {str(e)}")
        raise Exception(f"Error publishing to WordPress: {str(e)}")

@app.route('/api/process', methods=['POST'])
def process_file():
    try:
        logger.debug("Received process request")
        if 'file' not in request.files:
            logger.error("No file provided in request")
            return jsonify({'error': 'No file provided'}), 400
            
        file = request.files['file']
        action = request.form.get('action')
        title = request.form.get('title')
        
        logger.debug(f"Processing request - Action: {action}, Title: {title}")
        
        if not file.filename.endswith('.xlsx'):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Only Excel (.xlsx) files are supported'}), 400
            
        # Save the file temporarily
        temp_path = f"temp_{int(time.time())}.xlsx"
        logger.debug(f"Saving temporary file: {temp_path}")
        file.save(temp_path)
        
        try:
            if action == 'convert':
                logger.debug("Converting file to HTML")
                # Convert Excel to HTML
                html_content = convert_excel_to_html(temp_path)
                return jsonify({'html': html_content})
                
            elif action == 'publish':
                if not title:
                    logger.error("No title provided for publishing")
                    return jsonify({'error': 'Title is required for publishing'}), 400
                    
                logger.debug("Converting and publishing file")
                # Convert Excel to HTML
                html_content = convert_excel_to_html(temp_path)
                
                # Publish to WordPress
                page_url = publish_to_wordpress(title, html_content)
                return jsonify({'url': page_url})
                
            else:
                logger.error(f"Invalid action: {action}")
                return jsonify({'error': 'Invalid action'}), 400
                
        finally:
            # Clean up temporary file
            if os.path.exists(temp_path):
                logger.debug(f"Cleaning up temporary file: {temp_path}")
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Error processing request: {str(e)}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    logger.info("Starting Flask server")
    app.run(debug=True) 