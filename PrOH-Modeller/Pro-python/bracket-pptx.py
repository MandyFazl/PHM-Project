import csv
import os
import logging
import re
from subprocess import run
import sys

# Create a new logger instance for bracket.py
logger = logging.getLogger('bracket_logger')
logger.setLevel(logging.INFO)

# Create a file handler to output logs to a file
file_handler = logging.FileHandler('bracket.log')
file_handler.setLevel(logging.INFO)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
file_handler.setFormatter(formatter)
logger.addHandler(file_handler)

try:
    # Input CSV file name
    filename_with_identifier = sys.argv[1]
    filename_without_extension = os.path.splitext(filename_with_identifier)[0]
    file_path = os.path.join(filename_with_identifier)

    output_file = os.path.join(filename_without_extension + '_SubBubbles.csv')

    cell_with_parentheses = []

    with open(file_path, 'r', newline='') as csv_input:
        reader = csv.reader(csv_input)
        for row in reader:
            for cell in row:
                if '(' in cell or ')' in cell:
                    cell_with_parentheses.append(cell)

    with open(output_file, 'w', newline='') as csv_output:
        writer = csv.writer(csv_output)
        writer.writerow(['Cell with Parentheses'])
        writer.writerows([[cell] for cell in cell_with_parentheses])

    logger.info(f"Cells with parentheses have been saved to '{output_file}'")

    # Call the Bracket-Sketch-unique.py script to generate the PowerPoint presentation
    script_path = os.path.join(os.path.dirname(os.path.realpath(__file__)), 'Bracket-Sketch-unique.py')
    run(['python3', script_path, output_file])

except Exception as e:
    logger.error(f'An error occurred: {str(e)}')
