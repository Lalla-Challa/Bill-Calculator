# AI Utility Bill Automation Agent

This project provides an AI-powered automation agent that processes utility bill images, extracts key information (Name, Account Number, Due Date, Total Amount Due) using GPT-4 Vision, stores the extracted data in an Excel sheet, and calculates the total payable amount.

## Features

- **Image Analysis**: Utilizes GPT-4 Vision to analyze utility bill images.
- **Data Extraction**: Extracts Customer Name, Account Number, Due Date, and Total Amount Due.
- **Excel Export**: Stores extracted data in a structured Excel spreadsheet.
- **Total Calculation**: Calculates the sum of all 'Total Amount Due' from the processed bills.

## Prerequisites

Before you begin, ensure you have the following installed:

- Python 3.8+
- An OpenAI API key with access to GPT-4 Vision.

## Setup

1.  **Clone the repository (if you haven't already):**

    ```bash
    git clone https://github.com/Lalla-Challa/Bill-Calculator.git
    cd Bill-Calculator
    ```

2.  **Create a virtual environment (recommended):**

    ```bash
    python -m venv venv
    ```

3.  **Activate the virtual environment:**

    -   **On Windows:**

        ```bash
        .\venv\Scripts\activate
        ```

    -   **On macOS/Linux:**

        ```bash
        source venv/bin/activate
        ```

4.  **Install the required dependencies:**

    ```bash
    pip install -r requirements.txt
    ```

### API Key Configuration

1.  **OpenAI API Key**: Obtain your OpenAI API key from the [OpenAI platform](https://platform.openai.com/).
2.  **API Key Configuration**: The application will prompt you for your OpenAI API key if it's not found. You can also enter it in the GUI and click "Save API Key" to store it persistently in a `.env` file.
    *Note: The `.env` file is excluded from version control to protect your API key.*

## Usage

1.  **Run the Application**: Execute the `main.py` script:

    ```bash
    python main.py
    ```

2.  **API Key**: The application will attempt to load your API key from a `.env` file. If not found, enter your OpenAI API key into the designated field in the GUI and click "Save API Key" for future use.

3.  **Select Bill Images**: Click the "Select Bill Images" button to choose one or more utility bill image files (PNG, JPG, JPEG, GIF, BMP, TIFF).

4.  **Process Bills**: Click the "Process Bills" button to start the analysis. The application will display progress and any errors in the text area.

5.  **View and Save Output**: Once processing is complete, a link to the temporary Excel file will appear in the output area. You can click this link to open the Excel file directly. To save the file permanently, click the "Save Excel As..." button and choose your desired location.

## Output

After successful execution, a temporary Excel file will be generated. This file will contain:

-   A sheet named "Bill Details" with columns for Filename, Customer Name, Account Number, Due Date, Total Amount Due, Payable Within Due Date, and Payable After Due Date for each processed bill.
-   Summary rows at the bottom showing the "Total Payable Amount", "Total Payable Within Due Date", and "Total Payable After Due Date" across all bills.

## Error Handling

-   If the `OPENAI_API_KEY` environment variable is not set, the script will exit with an error message.
-   If an image cannot be processed by GPT-4 Vision, it will be skipped, and a message will be printed to the console.
-   If no images are found or no data is extracted, the Excel file will not be created, and a message will be displayed.
-   Temporary Excel files are automatically cleaned up when the application exits.

## Contributing

Feel free to fork the repository, open issues, or submit pull requests if you have suggestions or improvements.