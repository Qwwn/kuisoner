# KL Kuesioner Data Processing and Reporting

This project is designed to process survey data (KL Kuesioner) and generate a Word document containing tables and pie charts for each course.

## Installation

1. Clone the repository:

    ```bash
    git clone https://github.com/qwwn/kuisoner.git
    cd kuisoner
    ```

2. Install the required packages using pip:

    ```bash
    pip install -r requirements.txt
    ```

## Usage

1. Place your KL Kuesioner CSV file in the same directory as the script.

2. Run the script:

    ```bash
    python app.py
    ```

3. The generated Word document will be saved in the same directory with the name format: `KL_Kuesioner_YYYYMM_LecturerName_STMT.docx`.

## CSV File Format

Ensure that your KL Kuesioner CSV file follows the required format. The script expects columns named:
- 'Mata Kuliah'
- 'Pertanyaan'
- 'Sangat Setuju'
- 'Setuju'
- 'Tidak Setuju'
- 'Sangat Tidak Setuju'
