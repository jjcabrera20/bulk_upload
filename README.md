# KoboToolbox Bulk Upload Script

Uploads data from an Excel file to any KoboToolbox server.  
Uses **API v2** to fetch form metadata and **API v1** to submit records.

## Requirements

- Python 3.8+
- `pandas`, `requests`, `python-dotenv`

```bash
pip install pandas requests python-dotenv
```

## Configuration

All parameters live in a `.env` file at the project root:

| Variable | Description |
|---|---|
| `API_KEY` | Your KoboToolbox API token (Account → Security → API key) |
| `KOBO_URL` | Base URL of the KoboToolbox server |
| `ASSET_UID` | UID of the target form (alphanumeric, found in the form URL) |
| `EXCEL_FILE` | Path to the Excel file to upload (default: `data.xlsx`) |

The script automatically derives the submission server URL from `KOBO_URL`
(replaces `kf.` with `kc.` for public KoboToolbox; no change for self-hosted servers).

### Public KoboToolbox

```
API_KEY=your_api_token_here
KOBO_URL=https://kf.kobotoolbox.org
ASSET_UID=aXXXXXXXXXXXXXXXXXXXX
EXCEL_FILE=data.xlsx
```

### EU server

```
KOBO_URL=https://eu.kobotoolbox.org
```

### Self-hosted / custom server

```
KOBO_URL=https://your-kobo-server.org
```

## Usage

1. Fill in `.env` with your credentials and form details.
2. Prepare an Excel file whose column headers match either the **field names** or **field labels** of the KoboToolbox form.
3. Run:

```bash
python upload_generic.py
```

The script will:
1. Fetch the form structure via API v2.
2. Display available form fields.
3. Auto-map Excel columns to form fields (by name, then by label).
4. Ask for confirmation before uploading.
5. Submit each row as JSON via API v1 and print a success/failure summary.

## Column Mapping

- Exact match on field **name** is tried first.
- If no name match, the column is matched against the field **label**.
- Unmapped columns are listed as warnings and skipped.

## API Reference

- [KoboToolbox API v2](https://support.kobotoolbox.org/api.html)
- [API v1 → v2 Migration Guide](https://support.kobotoolbox.org/migrating_api.html)
- [Interactive API v2 docs](https://kf.kobotoolbox.org/api/v2/docs/)
