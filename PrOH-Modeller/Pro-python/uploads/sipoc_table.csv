from flask import Flask, render_template, request, send_file
import subprocess


app = Flask(__name__)
 

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/uploadfile', methods=['POST'])
def upload():
    if 'file' not in request.files:
        return "No file part"

    file = request.files['file']

    if file.filename == '':
        return "No selected file"

    file.save('uploads/' + file.filename)  # Save the uploaded file

    # Call your Python script with the uploaded file as an argument
    subprocess.run(['python', 'sipoc-to-pptx-4-Flask.py', 'uploads/' + file.filename])

    return "File uploaded and processed successfully"

if __name__ == '__main__':
    app.run(debug=True)





