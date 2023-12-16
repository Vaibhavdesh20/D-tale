from flask import Flask, request, jsonify,render_template
import pandas as pd
import dtale

app = Flask(__name__)

@app.route('/')
def index():
    return render_template('upload_excel.html')


@app.route('/upload_excel', methods=['POST'])
def handle_upload():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'})
    
    file = request.files['file']
    save_location = request.form.get('save_location', '.')  # Get save location or default to current directory
    
    if file.filename == '':
        return jsonify({'error': 'No selected file'})
    
    if file:
        try:
            # Save uploaded Excel file
            file_path = os.path.join(save_location, file.filename)
            file.save(file_path)
            
            # Convert Excel to CSV
            csv_data = convert_excel_to_csv(file_path)
            
            # Generate D-Tale report
            df = pd.read_csv(pd.compat.StringIO(csv_data))
            dtale_app = dtale.show(df)
            dtale_app.open_browser()
            dtale_report = dtale_app.build_url()
            
            return jsonify({'success': 'File converted and D-Tale report generated', 'report_url': dtale_report})
        except Exception as e:
            return jsonify({'error': str(e)})
    else:
        return jsonify({'error': 'File format not supported'})

def convert_excel_to_csv(excel_file):
    # Read Excel file into pandas DataFrame
    df = pd.read_excel(excel_file)
    # Convert DataFrame to CSV string
    csv_data = df.to_csv(index=False)
    return csv_data


if __name__ == '__main__':
    app.run(debug=True)
