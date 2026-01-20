# KoboToolbox Data Upload Script

This script uploads data from an Excel file to a KoboToolbox server.

## Requirements

*   Python 3.6+
*   `pandas` library
*   `requests` library

## Installation

1.  Install the required libraries:

    ```bash
    pip install pandas requests
    ```

## Usage

1.  Create an Excel file named `data.xlsx` with the following column headers:

    *   `AMIE`
    *   `Nombre Institución`
    *   `Zona`
    *   `Admin1`
    *   `Cod Admin1`
    *   `Admin2`
    *   `Cod Admin2`

2.  Update the `kobo_upload.py` script with your KoboToolbox credentials:

    *   `KOBOTOOLBOX_URL`
    *   `FORM_ID`
    *   `USERNAME`
    *   `PASSWORD`

3.  Run the script:

    ```bash
    python kobo_upload.py
    ```

## Error Handling

The script includes basic error handling and logging. If any errors occur during the data upload process, they will be printed to the console.