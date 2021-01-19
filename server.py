from flask import Flask, render_template, request, url_for, redirect
from excel_parser import ExcelParser

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('homepage.html')

@app.route('/done', methods=['POST'])
def done():
    file_name = request.form["file_name"]
    output_name = request.form["output_name"]
    threshold = request.form["threshold"]

    # if user wants output file to replace input file
    if output_name == '0':
        output_name = file_name

    parser = ExcelParser(file_name, output_name, int(threshold))
    parser.main()

    return render_template('done.html')

if __name__ == '__main__':
    app.run(host = '0.0.0.0', port = 3000, debug=True)
