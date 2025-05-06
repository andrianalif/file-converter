from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import os
import time
import logging
import base64
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Border, Side, Alignment
import warnings
import json
from datetime import datetime
from typing import List, Dict, Any

# Suppress the WMF image warning
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

app = Flask(__name__)
CORS(app)

# WordPress API credentials, change this to your own credentials
WP_USER = "alif.adrian"
WP_APP_PASSWORD = "xBa8 zJjb UDhK oXcP NeqR 7MDm"
WP_API_URL = "https://frontier.vstecs.ai/wp-json/wp/v2/pages"

logger.info("WordPress credentials loaded successfully")

def is_cell_bold(cell):
    """Helper function to safely check if a cell's font is bold"""
    try:
        if not cell or not hasattr(cell, 'font'):
            return False
        return cell.font.bold if cell.font else False
    except Exception as e:
        logger.debug(f"Error checking bold: {str(e)}")
        return False

def get_cell_value_safe(cell):
    """Helper function to safely get cell value"""
    try:
        if not cell:
            return ''
        return str(cell.value) if cell.value is not None else ''
    except Exception as e:
        logger.debug(f"Error getting cell value: {str(e)}")
        return ''

def get_cell_number_format(cell):
    """Helper function to safely get cell number format"""
    try:
        if not cell:
            return ''
        return str(cell.number_format) if cell.number_format else ''
    except Exception as e:
        logger.debug(f"Error getting number format: {str(e)}")
        return ''

def convert_to_rag_format(table_data: List[Dict[str, Any]], use_ollama: bool = True) -> str:
    """
    Convert table data to RAG format using OLLAMA with Gemma 3:27b or fallback to simple conversion
    """
    try:
        if use_ollama:
            try:
                # Prepare the prompt
                prompt = f"""
                Convert each row of the input table into exactly this format:
                "{{Item}}" is a "{{Description}}" that costs "{{Price}}"

                Rules:
                1. Infer columns by content:
                   - "Item" = Product code, model number, MTM, Series, or SKU (including those with #)
                   - "Description" = Combine all relevant product details into a comprehensive description
                   - "Price" = Field with currency/numeric value (USD, SGD, SRP, Disti Price, Reseller Price)
                2. For Description, include ALL of the following in order:
                   a. Product Type and Model:
                      - Product family (e.g., Desktop, Laptop, Workstation)
                      - Model name and generation
                      - Platform type (e.g., Intel SoC Platform)
                   
                   b. Processor Details:
                      - Full processor name (e.g., Intel® Core™ i7-13620H)
                      - Core configuration (e.g., 10C (6P + 4E) / 16T)
                      - Clock speeds for all core types (e.g., P-core 2.4 / 4.9GHz, E-core 1.8 / 3.6GHz)
                      - Cache size (e.g., 24MB)
                   
                   c. Memory and Storage:
                      - RAM type, size, and speed (e.g., 16GB SO-DIMM DDR5-5200)
                      - Storage type, size, and interface (e.g., 512GB SSD M.2 2280 PCIe® 4.0x4 NVMe®)
                   
                   d. Graphics:
                      - Graphics type and model (e.g., Integrated Intel® UHD Graphics)
                   
                   e. Connectivity:
                      - Wireless specifications (e.g., Intel® Wi-Fi® 6 AX203, 802.11ax 2x2 + BT5.2)
                      - Port types and configurations
                   
                   f. Input Devices:
                      - Keyboard type (e.g., USB Keyboard)
                      - Mouse type (e.g., USB Mouse)
                   
                   g. Operating System:
                      - OS version and edition (e.g., Windows® 11 Pro)
                      - Language
                      - Recovery media status (e.g., NO Recovery Media)
                   
                   h. Additional Features:
                      - Warranty information
                      - Special conditions or remarks
                      - Any other technical specifications
                3. For Price:
                   - Use USD if specified
                   - Use Suggested Retail Price (SRP) if available
                   - Fall back to Disti Price if SRP not available
                   - Include currency (USD/SGD) if specified
                   - Use "N/A" only for truly missing fields
                4. Additional Context:
                   - Include Bundle Type if available
                   - Include Product Family information
                   - Include Order Reason Code if present
                   - Include Special Conditions/UPC Code
                   - Include Additional Information
                5. Output one line per row - no headers, notes or explanations

                Input table data:
                {json.dumps(table_data, indent=2)}

                Example output:
                "TC30S5" is a "Desktop TC neo 30s Gen 5: Intel® SoC Platform - Intel® Core™ i7-13620H, 10C (6P + 4E) / 16T, P-core 2.4 / 4.9GHz, E-core 1.8 / 3.6GHz, 24MB - 16GB SO-DIMM DDR5-5200/ 512GB SSD M.2 2280 PCIe® 4.0x4 NVMe® - Graphics: Integrated Intel® UHD Graphics - Intel® Wi-Fi® 6 AX203, 802.11ax 2x2 + BT5.2 - USB Keyboard / Mouse - Windows® 11 Pro, English / NO Recovery Media" that costs "USD 1,354"
                """

                # Call OLLAMA API with timeout
                ollama_url = "http://172.18.0.221:7870/api/generate"
                response = requests.post(
                    ollama_url,
                    json={
                        "model": "gemma:3.27b",
                        "prompt": prompt,
                        "stream": False
                    },
                    timeout=30  # 30 seconds timeout
                )

                if response.status_code == 200:
                    return response.json().get("response", "")
                else:
                    logger.error(f"OLLAMA API error: {response.text}")
                    raise Exception(f"OLLAMA API error: {response.text}")
            except (requests.exceptions.ConnectionError, requests.exceptions.Timeout) as e:
                logger.warning(f"OLLAMA connection error, falling back to simple conversion: {str(e)}")
                use_ollama = False
            except Exception as e:
                logger.error(f"Error calling OLLAMA: {str(e)}")
                use_ollama = False

        # Fallback to simple conversion if OLLAMA is not available
        if not use_ollama:
            results = []
            for row in table_data:
                # Find the product code or model number
                item = "N/A"
                for key, value in row.items():
                    if isinstance(value, str) and value.strip():
                        key_lower = str(key).lower().replace('#', '')
                        if any(code_pattern in key_lower for code_pattern in ['code', 'model', 'part', 'sku', 'mtm', 'series']):
                            item = str(value).strip()
                            break
                        elif str(value).strip().isalnum() and len(str(value).strip()) < len(item):
                            item = str(value).strip()

                # Combine all fields for description
                description_parts = []
                
                # Handle Action field first if present
                if 'Action' in row and row['Action']:
                    description_parts.append(f"Action: {str(row['Action']).strip()}")
                
                # Handle other fields
                for key, value in row.items():
                    if isinstance(value, (str, int, float)) and str(value).strip():
                        key_lower = str(key).lower().replace('#', '')
                        if key_lower not in ['code', 'model', 'part', 'sku', 'mtm', 'series', 'price', 'srp', 'disti', 'reseller', 'usd', 'action']:
                            description_parts.append(f"{key}: {str(value).strip()}")
                
                description = " - ".join(description_parts) if description_parts else "N/A"

                # Find price field with priority
                price = "N/A"
                price_fields = [
                    'USD',
                    'Suggested Retail Price (SRP)',
                    'SRP',
                    'Disti Price',
                    'Reseller Price',
                    'Price'
                ]
                
                for field in price_fields:
                    if field in row:
                        value = row[field]
                        if isinstance(value, (int, float)):
                            if field == 'USD':
                                price = f"USD {value:,.2f}"
                            else:
                                price = f"SGD {value:,.2f}"
                            break
                        elif isinstance(value, str) and value.strip():
                            price = value.strip()
                            break

                results.append(f'"{item}" is a "{description}" that costs "{price}"')

            return "\n".join(results)

    except Exception as e:
        logger.error(f"Error in convert_to_rag_format: {str(e)}")
        return f"Error: {str(e)}"

def convert_excel_to_html(file_path):
    try:
        logger.debug(f"Loading workbook: {file_path}")
        wb = load_workbook(file_path, data_only=True)
        
        current_time = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
        
        # Initialize context storage
        product_contexts = []
        product_groups = {}
        current_category = None
        current_product_name = None
        
        # Add predefined product contexts for known models
        predefined_contexts = [
            {
                'category': 'Laptops',
                'product_name': 'Lenovo ThinkPad E14 Gen 6',
                'product_number': '21M7003KSG',
                'description': 'Lenovo ThinkPad E14 Gen 6 (14") Business Laptop with Intel Core i5-1335U, 16GB RAM, 512GB SSD, Windows 11 Pro, MIL-STD-810G tested, FHD (1920x1080) IPS display, Fingerprint Reader, Backlit Keyboard, 1 Year Warranty',
                'price': 1396,
                'subcategory': 'Business Laptops',
                'sheet_name': 'Laptops',
                'metadata': {
                    'is_subscription': False,
                    'is_service': False,
                    'specifications': {
                        'processor': 'Intel Core i5-1335U',
                        'ram': '16GB',
                        'storage': '512GB SSD',
                        'os': 'Windows 11 Pro',
                        'display': '14" FHD (1920x1080) IPS',
                        'features': ['Fingerprint Reader', 'Backlit Keyboard', 'MIL-STD-810G tested'],
                        'warranty': '1 Year',
                        'ports': ['USB-C', 'HDMI', 'USB-A', 'Headphone/Mic Combo Jack']
                    }
                }
            }
        ]
        product_contexts.extend(predefined_contexts)
        
        # CSS styles as a separate string
        css_styles = """
        <style>
            @import url('https://fonts.googleapis.com/css2?family=Poppins:wght@300;400;500;600;700&display=swap');
            
            body { 
                font-family: 'Poppins', sans-serif; 
                line-height: 1.6; 
                color: #333; 
                margin: 0; 
                padding: 20px; 
                background-color: #f4f4f4; 
                min-height: 100vh;
                display: flex;
                flex-direction: column;
            }
            h1, h2 { color: #2c3e50; margin-bottom: 20px; }
            .category-header { background-color: #f8f9fa; padding: 15px; margin: 20px 0; border-left: 5px solid #2c3e50; font-size: 1.2em; font-weight: bold; }
            .main-title { font-size: 2em; text-align: center; margin-bottom: 30px; color: #2c3e50; }
            .subtitle { font-size: 1.5em; text-align: center; margin: 20px 0; color: #34495e; }
            table { width: 100%; border-collapse: collapse; margin: 20px 0; background-color: #fff; border: 1px solid #ddd; box-shadow: 0 2px 4px rgba(0,0,0,0.1); }
            th, td { border: 1px solid #ddd; padding: 12px 15px; text-align: left; vertical-align: middle; }
            th { background-color: #f8f9fa; font-weight: bold; color: #2c3e50; border-bottom: 2px solid #ddd; }
            td { border-bottom: 1px solid #ddd; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            tr:hover { background-color: #f5f5f5; }
            .empty-row { height: 20px; background-color: transparent; border: none; }
            .empty-row td { border: none; }
            .price-column { text-align: right; font-family: monospace; white-space: nowrap; }
            .product-number-column { font-family: monospace; white-space: nowrap; }
            .description-column { min-width: 300px; }
            .section-header { background-color: #f8f9fa; font-weight: bold; color: #2c3e50; }
            .section-header td { padding: 15px; font-size: 1.1em; }
            .metadata { text-align: center; color: #666; margin: 10px 0; font-style: italic; }
        </style>
        """
        
        html_content = f"""
        {css_styles}
        <h1 class="main-title">Price List</h1>
        <div class="metadata">
            <p>Last Updated: {current_time}</p>
        </div>
        """
        
        # Store table data for RAG conversion
        table_data = []
        
        for sheet in wb.sheetnames:
            logger.debug(f"Processing sheet: {sheet}")
            ws = wb[sheet]
            
            html_content += f'<h2 class="subtitle">{sheet}</h2>\n'
            html_content += '<div class="table-container">\n<table>\n'
            
            is_header_section = False
            headers = {}
            first_row = next(ws.rows)
            
            for idx, cell in enumerate(first_row):
                header_value = get_cell_value_safe(cell)
                if header_value:
                    headers[idx] = str(header_value).strip()
            
            for row_idx, row in enumerate(ws.rows):
                if all(get_cell_value_safe(cell) == '' for cell in row):
                    if not is_header_section:
                        html_content += '<tr class="empty-row"><td colspan="100%">&nbsp;</td></tr>\n'
                    continue
                
                first_cell = row[0]
                if is_cell_bold(first_cell) and get_cell_value_safe(first_cell):
                    current_category = get_cell_value_safe(first_cell)
                    current_product_name = current_category
                    is_header_section = True
                    html_content += f'<tr class="section-header"><td colspan="100%">{current_category}</td></tr>\n'
                    continue
                
                is_header_section = False
                html_content += "<tr>"
                
                row_data = {}
                for col_idx, cell in enumerate(row):
                    value = get_cell_value_safe(cell)
                    number_format = get_cell_number_format(cell)
                    
                    header = headers.get(col_idx, '')
                    if col_idx == 0:
                        css_class = 'product-number-column'
                    elif any(price_term in str(header).lower() for price_term in ['price', 'usd', '$']):
                        css_class = 'price-column'
                    elif col_idx == 1:
                        css_class = 'description-column'
                    else:
                        css_class = ''
                    
                    if value == '':
                        display_value = ''
                    elif isinstance(value, (int, float)):
                        if any(price_term in str(header).lower() for price_term in ['price', 'usd', '$']) or '$' in number_format:
                            display_value = f"${value:,.0f}"
                        elif ',' in number_format or '.' in number_format:
                            display_value = f"{value:,.0f}"
                        else:
                            display_value = str(value)
                    else:
                        display_value = str(value)
                    
                    if row_idx > 0 and header:
                        row_data[header] = value
                    
                    html_content += f'<td class="{css_class}">{display_value}</td>'
                
                html_content += "</tr>\n"
                
                if row_idx > 0 and row_data:
                    # Add row data for RAG conversion
                    table_data.append(row_data)
                    
                    product_number = row_data.get('Product Number', '')
                    description_value = row_data.get('Description', '')
                    description_str = str(description_value) if description_value is not None else ''
                    
                    price_value = None
                    for key in row_data:
                        if any(price_term in str(key).lower() for price_term in ['price', 'usd', '$']):
                            price_value = row_data[key]
                            break
                    
                    # Ambil warranty jika ada
                    warranty_value = row_data.get('Warranty', '')
                    
                    product_context = {
                        'category': current_category or sheet,
                        'product_name': current_product_name,
                        'product_number': product_number,
                        'description': description_str,
                        'price': price_value,
                        'warranty': warranty_value,  # Tambahkan field warranty
                        'subcategory': current_category,
                        'sheet_name': sheet,
                        'metadata': {
                            'is_subscription': 'subscription' in description_str.lower(),
                            'is_service': 'service' in description_str.lower(),
                        }
                    }
                    product_contexts.append(product_context)
                    
                    if current_product_name:
                        if current_product_name not in product_groups:
                            product_groups[current_product_name] = {
                                'product_numbers': [],
                                'descriptions': {},
                                'prices': {}
                            }
                        if product_number:
                            product_groups[current_product_name]['product_numbers'].append(product_number)
                            product_groups[current_product_name]['descriptions'][product_number] = description_str
                            product_groups[current_product_name]['prices'][product_number] = price_value
            
            html_content += "</table>\n</div>\n"
        
        # Convert table data to RAG format
        rag_content = convert_to_rag_format(table_data)
        
        # Add RAG content to the page
        html_content += f"""
        <div class="rag-content" style="margin-top: 2rem; padding: 1rem; background-color: #f8f9fa; border-radius: 4px;">
            <h3>RAG Format</h3>
            <pre style="white-space: pre-wrap; font-family: 'Poppins', monospace; font-size: 0.9em;">
{rag_content}
            </pre>
        </div>
        """
        
        # Add structured data
        structured_data = {
            "@context": "https://schema.org/",
            "@type": "ItemList",
            "itemListElement": [
                {
                    "@type": "Product",
                    "name": ctx['product_name'],
                    "productNumber": ctx['product_number'],
                    "description": ctx['description'],
                    "category": ctx['category'],
                    "warranty": ctx.get('warranty', ''),  # Tambahkan field warranty
                    "offers": {
                        "@type": "Offer",
                        "price": str(ctx['price']) if ctx['price'] is not None else '',
                        "priceCurrency": "USD"
                    }
                }
                for ctx in product_contexts
            ],
            "productGroups": product_groups
        }
        
        html_content += f"""
        <script type="application/ld+json">
            {json.dumps(structured_data)}
        </script>
        <div id="product-contexts" style="display: none;" data-context-version="1.0">
            {json.dumps({
                'products': product_contexts,
                'groups': product_groups,
                'common_questions': [
                    {
                        'question': 'What are the product numbers for {}?',
                        'type': 'product_numbers_by_name'
                    },
                    {
                        'question': 'What is the price of {}?',
                        'type': 'price_by_product_number'
                    },
                    {
                        'question': 'Can you describe {}?',
                        'type': 'description_by_product_number'
                    }
                ]
            })}
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

@app.route('/api/rag/convert', methods=['POST'])
def rag_convert():
    """
    Convert Excel data to RAG format with customizable options
    """
    try:
        if 'file' not in request.files:
            return jsonify({
                'status': 'error',
                'message': 'No file provided'
            }), 400
            
        file = request.files['file']
        if not file.filename.endswith('.xlsx'):
            return jsonify({
                'status': 'error',
                'message': 'Invalid file type. Please upload an Excel (.xlsx) file.'
            }), 400

        # Get optional parameters
        format_type = request.form.get('format', 'default')  # default, json, csv
        include_metadata = request.form.get('include_metadata', 'false').lower() == 'true'
        custom_prompt = request.form.get('prompt')
        use_ollama = request.form.get('use_ollama', 'true').lower() == 'true'
            
        # Save the file temporarily
        temp_path = f"temp_{int(time.time())}.xlsx"
        file.save(temp_path)
        
        try:
            # Convert to RAG format
            wb = load_workbook(temp_path, data_only=True)
            table_data = []
            metadata = {
                'total_sheets': len(wb.sheetnames),
                'sheets': [],
                'total_rows': 0,
                'total_columns': 0,
                'conversion_method': 'ollama' if use_ollama else 'simple'
            }
            
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                headers = {}
                first_row = next(ws.rows)
                
                for idx, cell in enumerate(first_row):
                    header_value = get_cell_value_safe(cell)
                    if header_value:
                        headers[idx] = str(header_value).strip()
                
                sheet_data = []
                for row_idx, row in enumerate(ws.rows):
                    if row_idx == 0:  # Skip header row
                        continue
                        
                    row_data = {}
                    for col_idx, cell in enumerate(row):
                        if col_idx in headers:
                            row_data[headers[col_idx]] = get_cell_value_safe(cell)
                    
                    if row_data:  # Only add non-empty rows
                        sheet_data.append(row_data)
                        table_data.append(row_data)
                
                metadata['sheets'].append({
                    'name': sheet,
                    'rows': len(sheet_data),
                    'columns': len(headers)
                })
                metadata['total_rows'] += len(sheet_data)
                metadata['total_columns'] = max(metadata['total_columns'], len(headers))
            
            # Convert to RAG format
            rag_content = convert_to_rag_format(table_data, use_ollama)
            
            # Format response based on requested format
            if format_type == 'json':
                # Convert RAG content to structured JSON
                items = []
                for line in rag_content.split('\n'):
                    if not line.strip():
                        continue
                    try:
                        item = line.split(' is a ')[0].strip('"')
                        description = line.split(' is a ')[1].split(' that costs ')[0].strip('"')
                        price = line.split(' that costs ')[1].strip('"')
                        items.append({
                            'item': item,
                            'description': description,
                            'price': price
                        })
                    except:
                        continue
                
                result = {
                    'items': items,
                    'metadata': metadata if include_metadata else None
                }
            elif format_type == 'csv':
                # Convert to CSV format
                csv_lines = ['Item,Description,Price']
                for line in rag_content.split('\n'):
                    if not line.strip():
                        continue
                    try:
                        item = line.split(' is a ')[0].strip('"')
                        description = line.split(' is a ')[1].split(' that costs ')[0].strip('"')
                        price = line.split(' that costs ')[1].strip('"')
                        csv_lines.append(f'"{item}","{description}","{price}"')
                    except:
                        continue
                result = '\n'.join(csv_lines)
            else:
                # Default format (text)
                result = {
                    'content': rag_content,
                    'metadata': metadata if include_metadata else None
                }
            
            return jsonify({
                'status': 'success',
                'format': format_type,
                'data': result
            })
            
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Error in RAG conversion: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

@app.route('/api/rag/validate', methods=['POST'])
def rag_validate():
    """
    Validate if the Excel file structure is suitable for RAG conversion
    """
    try:
        if 'file' not in request.files:
            return jsonify({
                'status': 'error',
                'message': 'No file provided'
            }), 400
            
        file = request.files['file']
        if not file.filename.endswith('.xlsx'):
            return jsonify({
                'status': 'error',
                'message': 'Invalid file type'
            }), 400
            
        # Save the file temporarily
        temp_path = f"temp_{int(time.time())}.xlsx"
        file.save(temp_path)
        
        try:
            wb = load_workbook(temp_path, data_only=True)
            validation_results = {
                'is_valid': True,
                'issues': [],
                'recommendations': []
            }
            
            # Check each sheet
            for sheet in wb.sheetnames:
                ws = wb[sheet]
                headers = {}
                first_row = next(ws.rows)
                
                # Check headers
                for idx, cell in enumerate(first_row):
                    header_value = get_cell_value_safe(cell)
                    if header_value:
                        headers[idx] = str(header_value).strip()
                
                if len(headers) < 3:
                    validation_results['is_valid'] = False
                    validation_results['issues'].append(f"Sheet '{sheet}' has less than 3 columns")
                    validation_results['recommendations'].append(f"Add more columns to sheet '{sheet}'")
                
                # Check for required column types
                has_item = False
                has_description = False
                has_price = False
                
                for header in headers.values():
                    header_lower = header.lower()
                    if any(term in header_lower for term in ['item', 'product', 'sku', 'code']):
                        has_item = True
                    if any(term in header_lower for term in ['description', 'detail', 'info']):
                        has_description = True
                    if any(term in header_lower for term in ['price', 'cost', 'amount', '$']):
                        has_price = True
                
                if not has_item:
                    validation_results['recommendations'].append(f"Add an 'Item' or 'Product' column to sheet '{sheet}'")
                if not has_description:
                    validation_results['recommendations'].append(f"Add a 'Description' column to sheet '{sheet}'")
                if not has_price:
                    validation_results['recommendations'].append(f"Add a 'Price' column to sheet '{sheet}'")
            
            return jsonify({
                'status': 'success',
                'validation': validation_results
            })
            
        finally:
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
    except Exception as e:
        logger.error(f"Error in RAG validation: {str(e)}")
        return jsonify({
            'status': 'error',
            'message': str(e)
        }), 500

if __name__ == '__main__':
    logger.info("Starting Flask server")
    app.run(debug=True) 