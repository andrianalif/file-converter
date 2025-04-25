from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import os
import time
import logging
import base64
from openpyxl import load_workbook
import warnings
import json
from datetime import datetime

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
        # Load workbook with data_only=True to get calculated values instead of formulas
        wb = load_workbook(file_path, data_only=True)
        
        # Get current datetime
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Initialize context storage
        product_contexts = []
        current_category = None
        
        html_content = f"""
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
            .category-header {
                background-color: #f8f9fa;
                padding: 15px;
                margin: 20px 0;
                border-left: 5px solid #2c3e50;
                font-size: 1.2em;
                font-weight: bold;
            }
            .main-title {
                font-size: 2em;
                text-align: center;
                margin-bottom: 30px;
                color: #2c3e50;
            }
            .subtitle {
                font-size: 1.5em;
                text-align: center;
                margin: 20px 0;
                color: #34495e;
            }
            table { 
                width: 100%; 
                border-collapse: collapse; 
                margin: 20px 0; 
                background-color: #fff; 
                border: 1px solid #ddd;
                box-shadow: 0 2px 4px rgba(0,0,0,0.1);
            }
            th, td { 
                border: 1px solid #ddd; 
                padding: 12px 15px; 
                text-align: left; 
                vertical-align: middle;
            }
            th { 
                background-color: #f8f9fa;
                font-weight: bold;
                color: #2c3e50;
                border-bottom: 2px solid #ddd;
            }
            td {
                border-bottom: 1px solid #ddd;
            }
            tr:nth-child(even) {
                background-color: #f9f9f9;
            }
            tr:hover {
                background-color: #f5f5f5;
            }
            .empty-row {
                height: 20px;
                background-color: transparent;
                border: none;
            }
            .empty-row td {
                border: none;
            }
            .price-column {
                text-align: right;
                font-family: monospace;
                white-space: nowrap;
            }
            .product-number-column {
                font-family: monospace;
                white-space: nowrap;
            }
            .description-column {
                min-width: 300px;
            }
            .section-header {
                background-color: #f8f9fa;
                font-weight: bold;
                color: #2c3e50;
            }
            .section-header td {
                padding: 15px;
                font-size: 1.1em;
            }
            .metadata {
                text-align: center;
                color: #666;
                margin: 10px 0;
                font-style: italic;
            }
        </style>
        
        <h1 class="main-title">Price List</h1>
        <div class="metadata">
            <p>Last Updated: {current_time}</p>
        </div>
        """
        
        for sheet in wb.sheetnames:
            ws = wb[sheet]
            
            # Add sheet title
            html_content += f'<h2 class="subtitle">{sheet}</h2>\n'
            html_content += '<div class="table-container">\n<table>\n'
            
            # Track if we're currently in a header section
            is_header_section = False
            
            # Get column headers and their indices
            headers = {}
            first_row = next(ws.rows)
            for idx, cell in enumerate(first_row):
                if cell.value:
                    headers[idx] = str(cell.value).strip()
            
            # Process rows
            for row_idx, row in enumerate(ws.iter_rows()):
                # Skip completely empty rows but add some spacing
                if all(cell.value is None for cell in row):
                    if not is_header_section:
                        html_content += '<tr class="empty-row"><td colspan="100%">&nbsp;</td></tr>\n'
                    continue
                
                # Check if this is a header/category row (bold text in first column)
                first_cell = row[0]
                is_bold = False
                if hasattr(first_cell, 'font') and first_cell.font is not None:
                    is_bold = first_cell.font.bold
                
                if is_bold and first_cell.value:
                    current_category = first_cell.value
                    is_header_section = True
                    html_content += f'<tr class="section-header"><td colspan="100%">{current_category}</td></tr>\n'
                    continue
                
                # Regular row
                is_header_section = False
                html_content += "<tr>"
                
                row_data = {}
                for col_idx, cell in enumerate(row):
                    # Get cell value and format
                    value = cell.value
                    number_format = cell.number_format if hasattr(cell, 'number_format') and cell.number_format else ''
                    
                    # Determine column type for styling
                    header = headers.get(col_idx, '')
                    if col_idx == 0:  # Product Number column
                        css_class = 'product-number-column'
                    elif any(price_term in str(header).lower() for price_term in ['price', 'usd', '$']):
                        css_class = 'price-column'
                    elif col_idx == 1:  # Description column
                        css_class = 'description-column'
                    else:
                        css_class = ''
                    
                    # Format cell value based on type and format
                    if value is None:
                        display_value = ''
                    elif isinstance(value, (int, float)):
                        if any(price_term in str(header).lower() for price_term in ['price', 'usd', '$']) or '$' in number_format:
                            # Format as currency
                            display_value = f"${value:,.0f}"
                        elif ',' in number_format or '.' in number_format:
                            # Format as number with thousand separator
                            display_value = f"{value:,.0f}"
                        else:
                            display_value = str(value)
                    else:
                        display_value = str(value)
                    
                    # Store data for context
                    if row_idx > 0 and header:  # Skip header row
                        row_data[header] = value
                    
                    html_content += f'<td class="{css_class}">{display_value}</td>'
                
                html_content += "</tr>\n"
                
                # Store context for this product row
                if row_idx > 0 and row_data:
                    description_value = row_data.get('Description', '')
                    description_str = str(description_value) if description_value is not None else ''
                    
                    # Get price value, checking multiple possible column names
                    price_value = None
                    for key in row_data:
                        if any(price_term in str(key).lower() for price_term in ['price', 'usd', '$']):
                            price_value = row_data[key]
                            break
                    
                    product_context = {
                        'category': current_category or sheet,
                        'product_number': row_data.get('Product Number', ''),
                        'description': description_str,
                        'price': price_value,
                        'subcategory': current_category,
                        'sheet_name': sheet,
                        'metadata': {
                            'is_subscription': 'subscription' in description_str.lower(),
                            'is_service': 'service' in description_str.lower(),
                        }
                    }
                    product_contexts.append(product_context)
            
            html_content += "</table>\n</div>\n"
        
        # Add structured data
        structured_data = {
            "@context": "https://schema.org/",
            "@type": "ItemList",
            "itemListElement": [
                {
                    "@type": "Product",
                    "name": ctx['product_number'],
                    "description": ctx['description'],
                    "category": ctx['category'],
                    "offers": {
                        "@type": "Offer",
                        "price": str(ctx['price']) if ctx['price'] is not None else '',
                        "priceCurrency": "USD"
                    }
                }
                for ctx in product_contexts
            ]
        }
        
        html_content += f"""
        <script type="application/ld+json">
            {json.dumps(structured_data)}
        </script>
        <div id="product-contexts" style="display: none;" data-context-version="1.0">
            {json.dumps(product_contexts)}
        </div>
        """
        
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
            return jsonify({
                'status': 'error',
                'message': 'No file provided',
                'ui': {
                    'type': 'error',
                    'title': 'Upload Failed',
                    'description': 'Please select an Excel file to upload.'
                }
            }), 400
            
        file = request.files['file']
        action = request.form.get('action')
        
        # Get filename without extension for default title
        filename = os.path.splitext(file.filename)[0]
        # Use provided title or default to filename
        title = request.form.get('title') or filename
        
        logger.debug(f"Processing request - Action: {action}, Title: {title}, Original filename: {file.filename}")
        
        if not file.filename.endswith('.xlsx'):
            return jsonify({
                'status': 'error',
                'message': 'Invalid file type',
                'ui': {
                    'type': 'error',
                    'title': 'Invalid File',
                    'description': 'Please upload an Excel (.xlsx) file only.'
                }
            }), 400
            
        # Save the file temporarily
        temp_path = f"temp_{int(time.time())}.xlsx"
        logger.debug(f"Saving temporary file: {temp_path}")
        file.save(temp_path)
        
        try:
            if action == 'convert':
                html_content = convert_excel_to_html(temp_path)
                return jsonify({
                    'status': 'success',
                    'html': html_content,
                    'defaultTitle': filename,
                    'ui': {
                        'type': 'success',
                        'title': 'Conversion Successful',
                        'description': 'Your Excel file has been converted to HTML format.'
                    }
                })
                
            elif action == 'publish':
                if not title:
                    return jsonify({
                        'status': 'error',
                        'message': 'Title is required',
                        'ui': {
                            'type': 'error',
                            'title': 'Missing Title',
                            'description': 'Please provide a title for the price list.'
                        }
                    }), 400
                    
                html_content = convert_excel_to_html(temp_path)
                page_url = publish_to_wordpress(title, html_content)
                
                return jsonify({
                    'status': 'success',
                    'url': page_url,
                    'defaultTitle': filename,
                    'ui': {
                        'type': 'success',
                        'title': 'Published Successfully',
                        'description': f'Your price list has been published. View it <a href="{page_url}" target="_blank">here</a>.'
                    }
                })
                
            else:
                return jsonify({
                    'status': 'error',
                    'message': 'Invalid action',
                    'ui': {
                        'type': 'error',
                        'title': 'Invalid Action',
                        'description': 'The requested action is not supported.'
                    }
                }), 400
                
        finally:
            # Clean up temporary file
            if os.path.exists(temp_path):
                logger.debug(f"Cleaning up temporary file: {temp_path}")
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Error processing request: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': str(e),
            'ui': {
                'type': 'error',
                'title': 'Processing Error',
                'description': f'An error occurred: {str(e)}'
            }
        }), 500

if __name__ == '__main__':
    logger.info("Starting Flask server")
    app.run(debug=True) 