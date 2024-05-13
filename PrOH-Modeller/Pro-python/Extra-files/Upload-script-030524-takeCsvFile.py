@app.route('/upload', methods=['POST'])
def upload():
    try:
        if 'file' not in request.files:
            return jsonify({'error': 'No file selected'}), 400

        file = request.files['file']
        if file.filename == '':
            return jsonify({'error': 'No file selected'}), 400

        # Generate a unique identifier
        unique_identifier = str(uuid.uuid4())[:8]  # Adjust the length of the unique code as needed

        # Get the filename and extension
        filename, extension = os.path.splitext(file.filename)

        # Append the unique identifier to the filename
        filename_with_identifier = f"{filename}_{unique_identifier}{extension}"

        # Save the file with the new filename
        file_path = os.path.join('uploads', filename_with_identifier)
        file.save(file_path)
        logging.info("CSV file saved successfully.")

        # Store the file path in session
        session['file_path'] = file_path

        logging.info(f"Uploaded filename: {file.filename}")
        logging.info(f"Uploaded filepath: {file_path}")

        # Construct the filename without the extension
        filename_without_extension = os.path.splitext(filename_with_identifier)[0]
        logging.info(f"Filename without extension: {filename_without_extension}")

        # Call your python3 scripts with the uploaded file as an argument
        logging.info("calling your python3 scripts.")
        subprocess.run(['python', 'Sipoc_to_pptx.py', file_path])
        logging.info("Sipoc_to_pptx.py executed successfully.")
        subprocess.run(['python', 'Non-Core-Process-Statement.py', file_path])
        logging.info("Non_Core_Process_Statement.py executed successfully.")
        subprocess.run(['python', 'Seperate_verbs.py', file_path])
        logging.info("Seperate_verbs.py executed successfully.")
        subprocess.run(['python', 'Bracket.py', file_path])
        logging.info("Bracket.py executed successfully.")
        subprocess.run(['python', 'Decision-Bubbles.py', file_path])
        logging.info("Decision_Bubbles.py executed successfully.")
        subprocess.run(['python', 'RunAll.py', file_path])
        logging.info("RunAll.py executed successfully.")

        return jsonify({'message': 'Upload successful'}), 200
    except Exception as e:
        # Log any exceptions that occur
        logging.error(f'An error occurred: {str(e)}')
        return jsonify({'error': 'An error occurred while processing the file'}), 500