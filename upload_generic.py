import copy
import xml.etree.ElementTree as ET

import pandas as pd
import requests
import dotenv
import os
import uuid

dotenv.load_dotenv(override=True)

API_TOKEN  = os.getenv('API_KEY')
KOBO_URL   = os.getenv('KOBO_URL')
ASSET_UID  = os.getenv('ASSET_UID')
EXCEL_FILE = os.getenv('EXCEL_FILE', 'data.xlsx')

# For public KoboToolbox, submissions go to kc.* instead of kf.*
# For self-hosted servers both are on the same domain, so this is a no-op.
KOBO_KC_URL = KOBO_URL.replace('//kf.', '//kc.')


def get_form_structure():
    """Retrieve form structure from KoboToolbox API v2."""
    url = f"{KOBO_URL}/api/v2/assets/{ASSET_UID}/"
    headers = {"Authorization": f"Token {API_TOKEN}"}

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Error fetching form: {response.status_code}\n{response.text}")
        return None

    data = response.json()
    survey = data.get('content', {}).get('survey', [])

    skip_types = {'start', 'end', 'note', 'calculate',
                  'begin_group', 'end_group', 'begin_repeat', 'end_repeat'}

    fields = [
        {
            'name': item.get('name', ''),
            'type': item.get('type', ''),
            'label': (item.get('label') or [''])[0] if isinstance(item.get('label'), list) else item.get('label', ''),
        }
        for item in survey
        if item.get('type', '') not in skip_types and item.get('name')
    ]

    owner_url = data.get('owner', '') or ''
    username = owner_url.rstrip('/').rsplit('/', 1)[-1]

    print(f"\n✓ Form loaded: {data.get('name', 'Unknown')}")
    print(f"  UID: {data.get('uid')}  |  Version: {data.get('version_id')}  |  Fields: {len(fields)}")

    return {
        'id': data.get('uid'),
        'id_string': data.get('uid'),
        'title': data.get('name'),
        'fields': fields,
        'username': username,
    }


def fetch_instance_template(asset_uid):
    """Fetch the XForm XML definition and extract an empty default-instance
    element tree, stripped of its XForms namespace, to use as a submission
    template (preserves the form's group nesting)."""
    url = f"{KOBO_URL}/api/v2/assets/{asset_uid}.xml"
    headers = {"Authorization": f"Token {API_TOKEN}"}

    response = requests.get(url, headers=headers)
    if response.status_code != 200:
        print(f"Error fetching form XML: {response.status_code}\n{response.text}")
        return None

    root = ET.fromstring(response.content)
    for el in root.iter():
        if '}' in el.tag:
            el.tag = el.tag.split('}', 1)[1]

    instance = root.find('.//model/instance')
    if instance is None or len(instance) == 0:
        print("Error: could not find default instance in form XML.")
        return None

    return instance[0]


def read_excel_data(file_path):
    """Read data from Excel file."""
    try:
        df = pd.read_excel(file_path, dtype=str)
        print(f"\n✓ Read {len(df)} rows from {file_path}")
        return df
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return None


def map_columns_to_fields(df_columns, form_fields):
    """Map Excel columns to form field names (by name then by label)."""
    print("\n" + "=" * 60)
    print("COLUMN MAPPING")
    print("=" * 60)

    # Columns that are KoboToolbox export metadata — skip silently
    skip_columns = {
        'start', 'end', '_id', '_uuid', '_submission_time', '_validation_status',
        '_notes', '_status', '_submitted_by', '__version__', '_tags', '_index',
    }

    mapping = {}
    unmapped = []

    for col in df_columns:
        col_stripped = col.strip()

        # Skip blank column headers, KoboToolbox metadata, and select_multiple
        # choice columns (exported as "field/choice") and "Unnamed:" columns
        if (not col_stripped
                or col_stripped.lower() in skip_columns
                or col_stripped.startswith('Unnamed:')
                or '/' in col_stripped):
            continue

        col_lower = col_stripped.lower()
        match = next((f for f in form_fields if f['name'].lower() == col_lower), None)
        if not match:
            match = next((f for f in form_fields if f['label'].lower().strip() == col_lower), None)

        if match:
            mapping[col] = match['name']
            print(f"✓ '{col}' → '{match['name']}'")
        else:
            unmapped.append(col)

    if unmapped:
        print(f"\n⚠  Unmapped columns: {', '.join(unmapped)}")

    print("=" * 60)
    return mapping


def submit_row(username, instance_template, row_data: dict) -> tuple:
    """Submit a record via the OpenRosa submission endpoint (XML)."""
    record = copy.deepcopy(instance_template)

    for field_name, value in row_data.items():
        el = next(record.iter(field_name), None)
        if el is not None:
            el.text = value

    instance_id_el = next(record.iter('instanceID'), None)
    if instance_id_el is not None:
        instance_id_el.text = f"uuid:{uuid.uuid4()}"

    xml_body = ET.tostring(record, encoding='utf-8', xml_declaration=True)

    url = f"{KOBO_KC_URL}/{username}/submission"
    headers = {"Authorization": f"Token {API_TOKEN}"}
    files = {"xml_submission_file": ("submission.xml", xml_body, "text/xml")}

    response = requests.post(url, headers=headers, files=files)
    ok = response.status_code in (200, 201, 202)
    return ok, "" if ok else f"{response.status_code}: {response.text}"


def main():
    print("=" * 60)
    print("KOBOTOOLBOX DATA UPLOADER")
    print("=" * 60)

    print("\nFetching form structure...")
    form = get_form_structure()
    if not form:
        print("\n✗ Failed to load form. Check API_KEY, KOBO_URL, and ASSET_UID in .env.")
        return

    print("\nFetching submission XML template...")
    instance_template = fetch_instance_template(ASSET_UID)
    if instance_template is None:
        print("\n✗ Failed to load submission template.")
        return

    print("\nForm fields:")
    for i, f in enumerate(form['fields'], 1):
        print(f"  {i}. {f['name']}  ({f['label']})  [{f['type']}]")

    df = read_excel_data(EXCEL_FILE)
    if df is None:
        return

    print(f"\nExcel columns: {list(df.columns)}")
    mapping = map_columns_to_fields(df.columns, form['fields'])

    if not mapping:
        print("\n✗ No columns mapped. Ensure column headers match form field names or labels.")
        return

    print(f"\n{len(mapping)} column(s) mapped.")
    if input("\nProceed with upload? (yes/no): ").strip().lower() != 'yes':
        print("Upload cancelled.")
        return

    success_count = error_count = 0

    print("\n" + "=" * 60)
    print("UPLOADING")
    print("=" * 60)

    for index, row in df.iterrows():
        row_data = {
            field_name: (str(row[col]) if pd.notna(row[col]) else "")
            for col, field_name in mapping.items()
        }

        print(f"Row {index + 1}/{len(df)}...", end=" ")
        ok, msg = submit_row(form['username'], instance_template, row_data)

        if ok:
            success_count += 1
            print("✓")
        else:
            error_count += 1
            print(f"✗ {msg}")
            if error_count == 1:
                if input("\nContinue with remaining rows? (yes/no): ").strip().lower() != 'yes':
                    break

    print("\n" + "=" * 60)
    print("SUMMARY")
    print("=" * 60)
    print(f"Total: {len(df)}  |  Success: {success_count}  |  Failed: {error_count}")

    if success_count > 0:
        print(f"\n✓ View submissions at: {KOBO_URL}/#/forms/{ASSET_UID}/data")


if __name__ == "__main__":
    main()
