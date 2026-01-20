import pandas as pd
import requests
from datetime import datetime
import uuid
from xml.etree.ElementTree import Element, SubElement, tostring
import dotenv
import os

# Load environment variables from the .env file (if present)
dotenv.load_dotenv()

# Access environment variables as if they came from the actual environment
SECRET_KEY = os.getenv('API_KEY')
KOBO_URL = os.getenv('KOBO_URL')
KOBO_API_URL= os.getenv('KOBO_API_URL')
# Configuration
KOBO_URL = KOBO_URL  # KoboCAT server
KOBO_API_URL = KOBO_API_URL  # or "https://kobo.humanitarianresponse.info"
API_TOKEN = SECRET_KEY
ASSET_UID = "ayXYdZFhnyGajCRGUeREG6"  # The form UID from KoboToolbox
EXCEL_FILE = r"data.xlsx"

def get_form_structure():
    """Retrieve form structure from KoboToolbox API"""
    headers = {
        "Authorization": f"Token {API_TOKEN}"
    }
    
    # Try different API endpoints
    api_urls = [
        f"{KOBO_API_URL}/api/v2/assets/{ASSET_UID}/",
        f"{KOBO_API_URL}/assets/{ASSET_UID}/",
        f"{KOBO_URL}/api/v1/forms/{ASSET_UID}/form.json"
    ]
    
    for url in api_urls:
        print(f"  Trying: {url}")
        try:
            response = requests.get(url, headers=headers)
            if response.status_code == 200:
                print(f"  ✓ Success with: {url}")
                break
        except:
            continue
    
    try:
        if response.status_code == 200:
            data = response.json()
            
            form_id = data.get('uid')
            form_title = data.get('name', 'Unknown Form')
            version = data.get('version_id', 'v1')
            
            # Get form fields from content
            content = data.get('content', {})
            survey = content.get('survey', [])
            
            # Extract field information
            fields = []
            for item in survey:
                field_type = item.get('type', '')
                field_name = item.get('name', '')
                field_label = item.get('label', [''])[0] if isinstance(item.get('label'), list) else item.get('label', '')
                
                # Skip metadata and non-input fields
                if field_type in ['start', 'end', 'note', 'calculate']:
                    continue
                
                # Skip groups and repeats for now
                if field_type in ['begin_group', 'end_group', 'begin_repeat', 'end_repeat']:
                    continue
                
                fields.append({
                    'name': field_name,
                    'type': field_type,
                    'label': field_label
                })
            
            print(f"\n✓ Form loaded: {form_title}")
            print(f"  Form ID: {form_id}")
            print(f"  Version: {version}")
            print(f"  Fields found: {len(fields)}")
            
            return {
                'id': form_id,
                'title': form_title,
                'version': version,
                'fields': fields,
                'survey': survey
            }
        else:
            print(f"Error fetching form: {response.status_code}")
            print(response.text)
            
            # Try alternative: get form from KoboCAT directly
            print("\n  Trying alternative method from KoboCAT...")
            alt_url = f"{KOBO_URL}/api/v1/forms/{ASSET_UID}/form.json"
            print(f"  URL: {alt_url}")
            
            alt_response = requests.get(alt_url, headers=headers)
            if alt_response.status_code == 200:
                print("  ✓ Success with alternative method!")
                form_data = alt_response.json()
                
                # Parse the alternative format
                form_id = form_data.get('id_string', ASSET_UID)
                form_title = form_data.get('title', 'Unknown Form')
                version = form_data.get('version', 'v1')
                
                children = form_data.get('children', [])
                fields = []
                
                for item in children:
                    if item.get('type') not in ['start', 'end', 'note', 'calculate', 'group']:
                        fields.append({
                            'name': item.get('name', ''),
                            'type': item.get('type', ''),
                            'label': item.get('label', '')
                        })
                
                print(f"\n✓ Form loaded: {form_title}")
                print(f"  Form ID: {form_id}")
                print(f"  Version: {version}")
                print(f"  Fields found: {len(fields)}")
                
                return {
                    'id': form_id,
                    'title': form_title,
                    'version': version,
                    'fields': fields,
                    'survey': children
                }
            else:
                print(f"  ✗ Alternative method also failed: {alt_response.status_code}")
                print(f"  {alt_response.text}")
                return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def read_excel_data(file_path):
    """Read data from Excel file"""
    try:
        df = pd.read_excel(file_path, dtype=str)
        print(f"\n✓ Successfully read {len(df)} rows from Excel file")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None

def map_columns_to_fields(df_columns, form_fields):
    """Map Excel columns to form fields"""
    print("\n" + "=" * 60)
    print("COLUMN MAPPING")
    print("=" * 60)
    
    mapping = {}
    unmapped_columns = []
    
    # Try to match columns to fields
    for col in df_columns:
        matched = False
        
        # Try exact match with field name
        for field in form_fields:
            if col.lower().strip() == field['name'].lower():
                mapping[col] = field['name']
                matched = True
                print(f"✓ '{col}' → '{field['name']}' ({field['label']})")
                break
        
        # Try match with field label
        if not matched:
            for field in form_fields:
                if col.lower().strip() == field['label'].lower().strip():
                    mapping[col] = field['name']
                    matched = True
                    print(f"✓ '{col}' → '{field['name']}' (matched by label)")
                    break
        
        if not matched:
            unmapped_columns.append(col)
    
    if unmapped_columns:
        print(f"\n⚠️  Unmapped columns: {', '.join(unmapped_columns)}")
    
    print("=" * 60)
    
    return mapping

def get_field_path(field_name, survey):
    """Get the full path of a field including groups"""
    path = []
    group_stack = []
    
    for item in survey:
        item_type = item.get('type', '')
        item_name = item.get('name', '')
        
        if item_type == 'begin_group':
            group_stack.append(item_name)
        elif item_type == 'end_group':
            if group_stack:
                group_stack.pop()
        elif item_name == field_name:
            path = group_stack + [field_name]
            break
    
    return '/'.join(path) if path else field_name

def create_xml_submission(row, form_structure, column_mapping):
    """Create an XML submission from a DataFrame row"""
    current_time = datetime.now().strftime("%Y-%m-%dT%H:%M:%S.%f")[:-3]
    instance_id = f"uuid:{uuid.uuid4()}"
    
    form_id = form_structure['id']
    survey = form_structure['survey']
    
    # Create root element
    root = Element(form_id, {
        'id': form_id,
        'version': form_structure['version']
    })
    
    # Track groups
    group_stack = [root]
    group_elements = {None: root}
    
    # Process survey items in order
    for item in survey:
        item_type = item.get('type', '')
        item_name = item.get('name', '')
        
        if item_type == 'begin_group':
            # Create group element
            parent = group_stack[-1]
            group_elem = SubElement(parent, item_name)
            group_stack.append(group_elem)
            group_elements[item_name] = group_elem
            
        elif item_type == 'end_group':
            if len(group_stack) > 1:
                group_stack.pop()
                
        elif item_type == 'start':
            elem = SubElement(group_stack[-1], 'start')
            elem.text = current_time
            
        elif item_type == 'end':
            elem = SubElement(group_stack[-1], 'end')
            elem.text = current_time
            
        elif item_type not in ['note', 'calculate', 'begin_repeat', 'end_repeat']:
            # Regular field
            parent = group_stack[-1]
            elem = SubElement(parent, item_name)
            
            # Find value from Excel
            value = ""
            for excel_col, field_name in column_mapping.items():
                if field_name == item_name:
                    value = row.get(excel_col, "")
                    if pd.notna(value):
                        value = str(value)
                    else:
                        value = ""
                    break
            
            print(f"Field: {item_name}, Value before setting: {value}, Type: {type(value)}")
            elem.text = value
    
    # Add meta
    meta = SubElement(root, 'meta')
    instanceID = SubElement(meta, 'instanceID')
    instanceID.text = instance_id
    
    # Convert to string
    xml_string = '<?xml version="1.0" encoding="UTF-8"?>\n'
    xml_string += tostring(root, encoding='unicode')
    
    return xml_string, instance_id

def submit_to_kobo_xml(xml_data, instance_id):
    """Submit XML data to KoboToolbox using OpenRosa protocol"""
    url = f"{KOBO_URL}/submission"
    
    headers = {
        "Authorization": f"Token {API_TOKEN}"
    }
    
    files = {
        'xml_submission_file': (f'{instance_id}.xml', xml_data, 'text/xml')
    }
    
    try:
        response = requests.post(url, headers=headers, files=files)
        
        if response.status_code in [200, 201, 202]:
            return True, "Success"
        else:
            return False, f"Error {response.status_code}: {response.text}"
    except Exception as e:
        return False, str(e)

def main():
    """Main function to process and upload data"""
    print("=" * 60)
    print("GENERIC KOBOTOOLBOX DATA UPLOADER")
    print("=" * 60)
    
    # Get form structure
    print("\nFetching form structure...")
    form_structure = get_form_structure()
    
    if not form_structure:
        print("\n✗ Failed to load form structure. Check your API token and ASSET_UID.")
        return
    
    # Display form fields
    print("\nForm fields:")
    for i, field in enumerate(form_structure['fields'], 1):
        print(f"  {i}. {field['name']} - {field['label']} ({field['type']})")
    
    # Print form fields for debugging
    print("\nForm Fields:")
    for field in form_structure['fields']:
        print(f"  Name: {field['name']}, Label: {field['label']}")
    
    # Read Excel data
    df = read_excel_data(EXCEL_FILE)
    if df is None:
        return
    
    print(f"\nExcel columns: {list(df.columns)}")
    
    # Map columns
    column_mapping = map_columns_to_fields(df.columns, form_structure['fields'])
    
    if not column_mapping:
        print("\n✗ No columns could be mapped to form fields.")
        print("Ensure Excel column names match either field names or labels from the form.")
        return
    
    # Confirm before proceeding
    print(f"\n{len(column_mapping)} columns mapped successfully.")
    proceed = input("\nProceed with upload? (yes/no): ").strip().lower()
    
    if proceed != 'yes':
        print("Upload cancelled.")
        return
    
    # Process each row
    success_count = 0
    error_count = 0
    
    print("\n" + "=" * 60)
    print("UPLOADING DATA")
    print("=" * 60)
    
    for index, row in df.iterrows():
        print(f"\nRow {index + 1}/{len(df)}...", end=" ")
        
        xml_data, instance_id = create_xml_submission(row, form_structure, column_mapping)
        
        if index == 0:
            print(f"\n\nFirst submission preview:\n{xml_data[:400]}...\n")
        
        success, message = submit_to_kobo_xml(xml_data, instance_id)
        
        if success:
            success_count += 1
            print("✓ Success")
        else:
            error_count += 1
            print(f"✗ Failed: {message}")
            
            if error_count == 1:
                retry = input("\nContinue with remaining rows? (yes/no): ").strip().lower()
                if retry != 'yes':
                    break
    
    # Summary
    print("\n" + "=" * 60)
    print("UPLOAD SUMMARY")
    print("=" * 60)
    print(f"Total records: {len(df)}")
    print(f"Successful: {success_count}")
    print(f"Failed: {error_count}")
    print("=" * 60)
    
    if success_count > 0:
        print(f"\n✓ View submissions at:")
        print(f"https://ee.kobotoolbox.org/#/forms/{form_structure['id']}/summary")

if __name__ == "__main__":
    main()