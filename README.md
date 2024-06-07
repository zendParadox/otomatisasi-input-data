# Input Data Barang

This script is used to input data barang into the system.

## Getting Started

### Prerequisites

- Selenium
- OpenPyXL
- ChromeDriver

### Installation

1. Install the required packages using pip:

```sh
    pip install selenium openpyxl
```

2. Download the ChromeDriver executable and add it to your system's PATH.

### Usage

1. Run the script using Python:

```sh
    python input-data-barang.py
```

### Configuration

You can configure the script by modifying the following variables:

- `wb`: The Excel workbook object
- `sheetRange`: The range of cells in the Excel sheet that contains the data
- `driver`: The Selenium WebDriver object
- `username_input`: The username input field element
- `password_input`: The password input field element
- `level_select`: The level select element
- `login_button`: The login button element
- `kode_barang_input`: The kode barang input field element
- `nama_barang_input`: The nama barang input field element
- `jenis_barang_input`: The jenis barang input field element
- `satuan_barang_input`: The satuan barang input field element
- `simpan_button`: The simpan button element

### Troubleshooting

If you encounter any issues, you can try the following:

- Check the ChromeDriver version and ensure it is compatible with your Chrome browser version.
- Check the Excel file format and ensure it is compatible with the OpenPyXL library.
- Check the script's configuration variables and ensure they are correctly set.

## License

This script is licensed under the MIT License.

## Acknowledgments

This script uses the following libraries:

- Selenium
- OpenPyXL
- ChromeDriver

Thank you to the developers and maintainers of these libraries for their hard work and dedication.
