@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        file.save('uploads/' + 'sipoc_table.csv')  
        logging.info("CSV file saved successfully.")


        # Call your Python script with the uploaded file as an argument
        subprocess.run(['python', 'sipoc-to-pptx-4-Flask.py', 'uploads/' + file.filename])
        logging.info("sipoc-to-pptx-4-Flask.py executed successfully.")
        subprocess.run(['python', 'Seperate_verbs.py', 'uploads/' + file.filename])
        logging.info("Seperate_verbs.py executed successfully.")

        return jsonify({'message': 'Upload successful'}), 200
    except Exception as e:
        # Log any exceptions that occur
        logging.error(f'An error occurred: {str(e)}')
        return jsonify({'error': 'An error occurred while processing the file'}), 500