# TukuToi Client Reports
A python application to generate insightful PDF summaries of timeEdition exported CSV Files

## Description
This script reads CSV data from the timeEdition application, processes it, and generates a PDF report. One of the primary challenges during the implementation was arranging the layout of the PDF using the reportlab library.

## Setup and Installation
1. Clone the repository:
    ```
    git clone https://github.com/TukuToi/client_reports.git
    ```
2. Navigate into the project directory:
    ```
    cd client_reports
    ```
3. Install the required packages:
    ```
    pip install -r requirements.txt
    ```

## Usage
- Replace the logo.png with your own logo.
- replace the (currently empty) `data.csv` file with an export from the timeEdition app, in CSV UTF-8 format.
- Run the script with:
```
python main.py
```
- find the generated PDF Report in the project folder.

## Contributing
Pull requests are welcome. For major changes, please open an issue first to discuss what you would like to change.

## License
[GPLv3](https://www.gnu.org/licenses/gpl-3.0.html)