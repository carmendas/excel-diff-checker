import glob
import os
import sys
import zipfile
import openpyxl
import pandas as pd
import pandas.io.formats.excel
import numpy as np
from flask import Flask, render_template, url_for, request, flash, redirect
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Get the current working directory
current_directory = os.getcwd()

# Replace backslashes with forward slashes
current_directory = current_directory.replace("\\", "/")

# Upload folder path
UPLOAD_FOLDER = current_directory + '/static/uploads'

allowed_extension = {'xlsx'}
app.config['SECRET_KEY'] = 'ExcelDiffCheck'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Remove default formatting in header when converting pandas DataFrame to Excel sheet
pd.io.formats.excel.ExcelFormatter.header_style = None


@app.route("/")
def home():
    """
        Function that render template
    """
    return render_template('upload-file.html')


def allowed_file(filename):
    return '.' in filename and \
        filename.rsplit('.', 1)[1].lower() in allowed_extension


@app.route('/', methods=['GET', 'POST'])
def upload_file():
    """
        Function that receives files sent by the form for processing
    """

    # Initialize the form_submitted variable as False for the initial load
    form_submitted = False

    if request.method == 'POST':
        # check if the post request has the file part
        if 'ref_file' and 'orig_file' not in request.files:
            flash('No file part')
            return redirect(request.url)

        reference_file = request.files['ref_file']
        original_file = request.files['orig_file']
        # If the user does not select a file, the browser submits an
        # empty file without a filename.
        if reference_file.filename == '' and original_file.filename == '':
            flash('No selected file')
            return redirect(request.url)

        if (reference_file and allowed_file(reference_file.filename)) and \
        (original_file and allowed_file(original_file.filename)):
            reference_file_name = secure_filename(reference_file.filename)
            original_file_name = secure_filename(original_file.filename)

        reference_file.save(os.path.join(app.config['UPLOAD_FOLDER'], reference_file_name))
        original_file.save(os.path.join(app.config['UPLOAD_FOLDER'], original_file_name))

        ref_require_cols = [int(request.form['ref_column'])]
        orig_require_cols = [int(request.form['orig_column'])]

        url = dataframe_comparison(reference_file_name, original_file_name, ref_require_cols, orig_require_cols)

        # Set a variable to indicate that the form has been submitted
        form_submitted = True

        return render_template('upload-file.html', form_submitted=form_submitted, zip_url=url)


def dataframe_comparison(reference_file_name, original_file_name, ref_require_cols, orig_require_cols):
    """
        Function to compare dataframe between the files sending by form
    """
    # Convert form file to numpy array
    reference_dataframe = pd.read_excel(UPLOAD_FOLDER + '/' + reference_file_name, usecols=ref_require_cols).to_numpy()
    original_dataframe = pd.read_excel(UPLOAD_FOLDER + '/' + original_file_name, usecols=orig_require_cols).to_numpy()
    # Check and verify values in files and return boolean array having same size as of original_dataframe
    dataframe = np.isin(original_dataframe, reference_dataframe)
    # Display the full value of numbers in files
    np.set_printoptions(threshold=sys.maxsize)

    # Extract found and not found values
    datas_found = np.where(dataframe == True)
    datas_not_found = np.where(dataframe == False)

    # Get original file
    original_file = pd.read_excel(UPLOAD_FOLDER + '/' + original_file_name)

    # Convert found and not found values to dataframes datas
    original_file_data_found = pd.DataFrame(original_file.iloc[datas_found[0]])
    original_file_data_not_found = pd.DataFrame(original_file.iloc[datas_not_found[0]])

    # Get original filename without the extension
    original_file_name_without_ext = os.path.splitext(os.path.basename(UPLOAD_FOLDER + '/' + original_file_name))[0]

    # Create new folder for generated files
    folder_path = UPLOAD_FOLDER + '/' + original_file_name_without_ext
    new_folder_path = create_diff_folder(folder_path)

    # Create Excel file for found and not found datas
    original_file_data_found.to_excel(new_folder_path + '/' + original_file_name_without_ext + '_found.xlsx', sheet_name='Sheet1', index=False)
    original_file_data_not_found.to_excel(new_folder_path + '/' + original_file_name_without_ext + '_not_found.xlsx', sheet_name='Sheet1', index=False)

    # Auto-adjust the new file column widths
    adjust_excel_file_column_to_column_content(new_folder_path + '/' + original_file_name_without_ext + '_found.xlsx')
    adjust_excel_file_column_to_column_content(new_folder_path + '/' + original_file_name_without_ext + '_not_found.xlsx')

    # Zip difference folder
    zip_url = zip_diff_result_folder(folder_path)

    return zip_url


def create_diff_folder(folder_path):
    """
        Function to create difference folder which contain found and not found file and return folder path
    """
    # Check if the folder exists
    if not os.path.exists(folder_path):
        # If it doesn't exist, create it
        os.makedirs(folder_path)
        return folder_path


def zip_diff_result_folder(folder_path):
    """
        Function to zip difference folder which contain found and not found file and return zip folder url
    """
    # Compress folder with zipfile
    with zipfile.ZipFile(folder_path + '.zip', 'w') as f:
        for file in glob.glob(folder_path + '/*'):
            f.write(file)
            zip_file_path = folder_path + '.zip'
            zip_file = os.path.basename(zip_file_path).split('/')[-1]
            zip_file_url = url_for('static', filename='uploads/'+zip_file)
            return zip_file_url


def adjust_excel_file_column_to_column_content(file_path):
    """
        Function to adjust found and not found Excel file column width
    """
    # Load the existing Excel file
    workbook = openpyxl.load_workbook(file_path)

    # Specify the sheet you want to adjust (replace 'Sheet1' with the actual sheet name)
    sheet = workbook['Sheet1']

    # Iterate through the columns and adjust the widths
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter  # Get the column letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except Exception as e:
                print(e)
        adjusted_width = (max_length + 2)
        sheet.column_dimensions[column_letter].width = adjusted_width

        # Save the modified Excel file
        workbook.save(file_path)
