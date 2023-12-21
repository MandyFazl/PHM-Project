import pandas as pd

def excel_to_html(input_excel, output_html):
    # Read Excel file into a Pandas DataFrame
    df = pd.read_excel(input_excel)

    # Convert DataFrame to HTML table
    html_table = df.to_html(index=False, classes='table table-bordered table-striped')

    # Write the HTML table to an output file
    with open(output_html, 'w') as f:
        f.write(html_table)

if __name__ == "__main__":
    # Replace 'input_excel.xlsx' with the path to your Excel file
    input_excel_file = 'Excel-format2.xlsx'

    # Replace 'output_table.html' with the desired output HTML file path
    output_html_file = 'output_table2.html'

    excel_to_html(input_excel_file, output_html_file)
