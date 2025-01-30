from flask import Flask, request, jsonify, send_file
import os
import pandas as pd
import logging
from werkzeug.utils import secure_filename

# Initialize Flask app
app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)

# Directory to store uploaded and processed files
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Allowed file extensions
ALLOWED_EXTENSIONS = {"xls", "xlsx"}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def process_excel(excel1_path, excel2_path):
    try:
        df1 = pd.read_excel(excel1_path)
        df2 = pd.read_excel(excel2_path)

        required_columns = ['Description', 'Material', 'Rate', 'C/kg']
        for col in required_columns[:2]:  # Only 'Description' and 'Material' are mandatory in Excel 2
            if col not in df2.columns:
                return None, f"Excel 2 must contain the column '{col}'"

        df2_filled = df2.copy()
        for index, row in df2.iterrows():
            if pd.isnull(row.get('Rate')) or pd.isnull(row.get('C/kg')):
                match = df1[(df1['Description'] == row['Description']) & (df1['Material'] == row['Material'])]
                if not match.empty:
                    if pd.isnull(row.get('Rate')) and 'Rate' in match.columns:
                        df2_filled.at[index, 'Rate'] = match.iloc[0]['Rate']
                    if pd.isnull(row.get('C/kg')) and 'C/kg' in match.columns:
                        df2_filled.at[index, 'C/kg'] = match.iloc[0]['C/kg']

        missing_data = df2_filled[df2_filled['Rate'].isnull() | df2_filled['C/kg'].isnull()]
        missing = not missing_data.empty

        output_file_name = secure_filename(os.path.basename(excel2_path).replace(".xlsx", "_updated.xlsx"))
        output_path = os.path.join(PROCESSED_FOLDER, output_file_name)
        df2_filled.to_excel(output_path, index=False)

        missing_data_path = None
        if missing:
            missing_file_name = secure_filename(os.path.basename(excel2_path).replace(".xlsx", "_missing.xlsx"))
            missing_data_path = os.path.join(PROCESSED_FOLDER, missing_file_name)
            missing_data.to_excel(missing_data_path, index=False)

        return output_path, missing_data_path
    except Exception as e:
        logging.error(f"Error processing Excel: {str(e)}")
        return None, str(e)

@app.route("/upload", methods=["POST"])
def upload_files():
    if "excel1" not in request.files or "excel2" not in request.files:
        return jsonify({"error": "Both Excel1 and Excel2 files are required."}), 400
    
    excel1 = request.files["excel1"]
    excel2 = request.files["excel2"]
    
    if not (excel1 and allowed_file(excel1.filename)) or not (excel2 and allowed_file(excel2.filename)):
        return jsonify({"error": "Invalid file format. Only .xls and .xlsx are allowed."}), 400
    
    excel1_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel1.filename))
    excel2_path = os.path.join(UPLOAD_FOLDER, secure_filename(excel2.filename))
    excel1.save(excel1_path)
    excel2.save(excel2_path)
    
    processed_path, missing_data_path = process_excel(excel1_path, excel2_path)
    if processed_path is None:
        return jsonify({"error": missing_data_path}), 500
    
    response = {"processed_file": processed_path}
    if missing_data_path:
        response["missing_data_file"] = missing_data_path
    
    return jsonify(response)

@app.route("/download/<filename>", methods=["GET"])
def download_file(filename):
    file_path = os.path.join(PROCESSED_FOLDER, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    return jsonify({"error": "File not found."}), 404

if __name__ == "__main__":
    port = int(os.environ.get("PORT", 8000))  # Automatically get Render's port
    app.run(host="0.0.0.0", port=port, debug=False)  # Debug should be False in production
