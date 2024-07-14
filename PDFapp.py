from flask import Flask, request, render_template_string, send_file, session
from flask_sqlalchemy import SQLAlchemy  # For SQLAlchemy integration with SQLite
import PyPDF2
import pandas as pd
import tempfile
import os

# Initialize Flask app
app = Flask(__name__)

# Configure SQLite database path
app.config['SQLALCHEMY_DATABASE_URI'] = 'sqlite:///' + os.path.join(os.path.dirname(__file__), 'database.db')
app.config['SQLALCHEMY_TRACK_MODIFICATIONS'] = False

# Initialize SQLAlchemy
db = SQLAlchemy(app)

# HTML template for the upload page
html_template = '''
<!doctype html>
<html>
    <head>
        <title>Upload PDF</title>
    </head>
    <body>
        <h1>Upload a PDF file</h1>
        <form action="/" method="post" enctype="multipart/form-data">
            <input type="file" name="file">
            <input type="submit" value="Upload">
        </form>
        {% if text %}
            <h2>Extracted Text:</h2>
            <pre>{{ text }}</pre>
            <br>
            <a href="/download">Download as XLS</a>
        {% endif %}
    </body>
</html>
'''

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    text = None
    if request.method == 'POST':
        file = request.files['file']
        if file:
            reader = PyPDF2.PdfReader(file.stream)
            text = ""
            for page_number in range(len(reader.pages)):
                page = reader.pages[page_number]
                text += page.extract_text() + "\n"
            session['extracted_text'] = text

    return render_template_string(html_template, text=text)

@app.route('/download')
def download_excel():
    text = session.get('extracted_text')
    if text:
        # Create a DataFrame from the extracted text
        df = pd.DataFrame({'Text': [text]})

        # Create a temporary file to save the Excel
        temp_file = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx')
        excel_file_path = temp_file.name

        # Export DataFrame to Excel
        df.to_excel(excel_file_path, index=False)

        # Serve the file for download
        return send_file(excel_file_path, as_attachment=True, download_name='extracted_text.xlsx')

    return "No text to download."


if __name__ == '__main__':
    app.secret_key = 'super_secret_key'
    app.run(debug=True, port=5003)