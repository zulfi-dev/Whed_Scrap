# WHED.net Scraper

This is a Python script for scraping data from the WHED.net website. It automates the process of extracting university information from the website and storing it in an Excel file.

## Description

The WHED.net Scraper uses Selenium WebDriver to navigate the website, interact with elements, and extract data. It retrieves university details such as name, country, address, staff count, student count, degree programs, and more. The scraped data is then saved to an Excel file for further analysis or processing.

## Features

- Scrapes university information from WHED.net
- Retrieves details such as name, country, address, staff count, student count, degree programs, and more
- Stores the scraped data in an Excel file
- Supports navigating through multiple result pages

## Requirements

- Python 3.x
- Selenium WebDriver
- pandas
- openpyxl

## Usage

1. Install the required dependencies using pip:
```bash
pip install Selenium pandas openpyxl

2. Download the WHED.net Scraper code from this repository.

3. Run the script:
```bash
python script.py

4. The scraped university data will be saved to an Excel file named `university_data.xlsx`.

## Contributions

Contributions to this project are welcome. If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.

## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
