# PDF Table Extraction to Excel (one sheet per table)

This project is designed to extract tables from PDF files and export them into an Excel file with multiple sheets (one sheet per table). It uses `PyMuPDF` for extracting tables and `pandas` for creating Excel files.

## Requirements

To run the script, you need to have the following Python packages installed:

- `fitz` (PyMuPDF)
- `pandas`
- `xlsxwriter`

You can install these dependencies by running the following command:

```bash
pip install PyMuPDF pandas xlsxwriter
```
## Usage

- **Upload the PDF file**: Place the PDF you want to process in the `uploads` folder.
- **Run the script**: The script will ask for the name of the PDF file (with the `.pdf` extension). Ensure that the file is located in the `uploads` folder.
- **Processing and Output**: The script will extract the tables from the PDF and save them as an Excel file in the `outputs` folder. Each table will be placed in a separate sheet in the Excel file.

## Functionality

- **Extract Tables**: The script uses PyMuPDF to open and extract tables from each page of the PDF.
- **Export to Excel**: The extracted tables are saved in an Excel file with one sheet per table using the pandas library.

## Notes

- The script assumes that the PDF contains tables in a standard format that PyMuPDF can detect. Complex table structures or poorly formatted tables may not be extracted properly.
- Make sure that the `uploads` and `outputs` directories exist. If the `outputs` directory does not exist, it will be created automatically.

## File Structure
```
📂 project-directory
 ├── 📂 uploads  # Place scanned PDFs here
 ├── 📂 outputs  # Processed files will be saved here
 ├── main.py  # Main script for processing PDFs
 ├── README.md  # Project documentation
 ├── LICENSE  # MIT License file
```

## Contributing
Feel free to contribute by submitting issues or pull requests.

## License
This project is licensed under the MIT License.

## Author
[Kurama-90](https://github.com/Kurama-90)
