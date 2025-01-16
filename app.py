from flask import Flask, request, send_file, render_template
import os
from resume_parser import Resume, EntityGenerator, json_to_excel
import tempfile

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    if 'resumes' not in request.files:
        return "No file part", 400

    files = request.files.getlist('resumes')

    # Check if filenames were provided
    if len(files) == 0:
        return "No files selected", 400

    # Process the files
    filenames = []
    all_responses = []

    for file in files:
        if file.filename == '':
            continue

        # Save the uploaded files to a temporary directory
        temp_filename = os.path.join(tempfile.gettempdir(), file.filename)
        file.save(temp_filename)
        filenames.append(temp_filename)

        # Process the resume
        resume = Resume(filename=temp_filename)
        try:
            response_news = resume.get()
        except Exception as e:
            return f"Error processing file '{file.filename}': {e}", 400

        helper = EntityGenerator(text=response_news)
        response = helper.get()
        all_responses.append(response)

    # Create an Excel file with the results
    if all_responses:
        # Create a temporary Excel file
        excel_filename = os.path.join(tempfile.gettempdir(), 'output.xlsx')
        json_to_excel(all_responses, filenames, excel_filename)

        # Send the Excel file to the user
        return send_file(excel_filename, as_attachment=True, download_name='resumes_entities.xlsx')

    return "No valid resumes were processed", 400

if __name__ == '__main__':
    app.run(debug=True)
