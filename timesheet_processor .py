from flask import Flask, request, jsonify
import pandas as pd
import json
from datetime import datetime
import os
from werkzeug.utils import secure_filename
import tempfile

app = Flask(__name__)

# Configure upload settings
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def parse_timesheet_data(file_path):
    """
    Parse the Excel timesheet and extract data in the required format
    """
    try:
        # Read the Excel file
        df = pd.read_excel(file_path, sheet_name=0, header=None)
        
        # Extract header information
        employee_name = ""
        sesa_id = ""
        reporting_manager_schneider = ""
        project_name = ""
        wissen_employee_id = ""
        reporting_manager_wissen = ""
        
        # Parse header section (first few rows)
        for i in range(min(10, len(df))):  # Check first 10 rows for header info
            for j in range(min(10, len(df.columns))):  # Check first 10 columns
                cell_value = str(df.iloc[i, j]).strip() if not pd.isna(df.iloc[i, j]) else ""
                
                # Extract employee information based on your timesheet format
                if "Employee Name:" in cell_value or (i == 1 and j == 1):  # Adjust based on actual position
                    try:
                        employee_name = str(df.iloc[i, j+1]) if j+1 < len(df.columns) else ""
                    except:
                        employee_name = ""
                
                if "SESA ID:" in cell_value or (i == 2 and j == 1):  # Adjust based on actual position
                    try:
                        sesa_id = str(df.iloc[i, j+1]) if j+1 < len(df.columns) else ""
                    except:
                        sesa_id = ""
                
                if "Project Name:" in cell_value:
                    try:
                        project_name = str(df.iloc[i, j+1]) if j+1 < len(df.columns) else ""
                    except:
                        project_name = ""
        
        # Find the timesheet data section (looking for date headers)
        timesheet_data = []
        date_row_index = -1
        
        # Look for the row containing dates
        for i in range(len(df)):
            row_has_dates = False
            for j in range(len(df.columns)):
                cell_value = str(df.iloc[i, j]).strip() if not pd.isna(df.iloc[i, j]) else ""
                # Look for date patterns like "Jan'25", "Feb'25", etc.
                if any(month in cell_value for month in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                                                        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']):
                    row_has_dates = True
                    break
            
            if row_has_dates:
                date_row_index = i
                break
        
        if date_row_index == -1:
            # Fallback: assume timesheet data starts from a specific row
            date_row_index = 5  # Adjust based on your template
        
        # Extract date headers
        date_headers = []
        month_year = ""
        
        for j in range(len(df.columns)):
            cell_value = str(df.iloc[date_row_index, j]).strip() if not pd.isna(df.iloc[date_row_index, j]) else ""
            if cell_value and cell_value != "nan":
                date_headers.append((j, cell_value))
                if any(month in cell_value for month in ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 
                                                        'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']):
                    month_year = cell_value
        
        # Extract timesheet values for each day
        for i in range(date_row_index + 1, len(df)):
            day_number = str(df.iloc[i, 0]).strip() if not pd.isna(df.iloc[i, 0]) else ""
            
            # Skip empty rows or non-numeric day values
            if not day_number or day_number == "nan":
                continue
            
            try:
                day_num = int(float(day_number))
            except (ValueError, TypeError):
                continue
            
            # Extract timesheet values for this day
            for col_index, date_header in date_headers:
                if col_index < len(df.columns):
                    timesheet_value = df.iloc[i, col_index]
                    
                    # Skip empty cells
                    if pd.isna(timesheet_value) or str(timesheet_value).strip() == "":
                        continue
                    
                    # Convert timesheet value to appropriate format
                    if isinstance(timesheet_value, (int, float)):
                        ts_value = float(timesheet_value)
                    else:
                        ts_str = str(timesheet_value).strip()
                        if ts_str and ts_str != "nan":
                            try:
                                ts_value = float(ts_str)
                            except ValueError:
                                ts_value = ts_str  # Keep as string if not numeric
                        else:
                            continue
                    
                    # Create the output record
                    record = {
                        "employee_name": employee_name,
                        "SESA_ID": sesa_id,
                        "reporting_manager-schneider": reporting_manager_schneider,
                        "project_name": project_name,
                        "wissen_employee_id": wissen_employee_id,
                        "reporting_manager-wissen": reporting_manager_wissen,
                        "Month": month_year,
                        "Date": f"{day_num:02d}",
                        "timesheet_value": ts_value
                    }
                    
                    timesheet_data.append(record)
        
        return timesheet_data
    
    except Exception as e:
        print(f"Error parsing timesheet: {str(e)}")
        return []

@app.route('/', methods=['GET'])
def home():
    return jsonify({
        "message": "Timesheet Processor API",
        "endpoints": {
            "/upload": "POST - Upload Excel timesheet file",
            "/health": "GET - Health check"
        }
    })

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy", "timestamp": datetime.now().isoformat()})

@app.route('/upload', methods=['POST'])
def upload_timesheet():
    try:
        # Check if file is present in request
        if 'file' not in request.files:
            return jsonify({"error": "No file provided"}), 400
        
        file = request.files['file']
        
        if file.filename == '':
            return jsonify({"error": "No file selected"}), 400
        
        if not allowed_file(file.filename):
            return jsonify({"error": "Invalid file type. Only .xls and .xlsx files are allowed"}), 400
        
        # Create a temporary file to save the uploaded file
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_file:
            file.save(temp_file.name)
            temp_file_path = temp_file.name
        
        try:
            # Process the timesheet
            timesheet_data = parse_timesheet_data(temp_file_path)
            
            if not timesheet_data:
                return jsonify({"error": "No valid timesheet data found in the file"}), 400
            
            # Return the processed data as JSON
            response_data = {
                "status": "success",
                "total_records": len(timesheet_data),
                "data": timesheet_data,
                "processed_at": datetime.now().isoformat()
            }
            
            return jsonify(response_data)
            
        finally:
            # Clean up temporary file
            if os.path.exists(temp_file_path):
                os.unlink(temp_file_path)
    
    except Exception as e:
        return jsonify({"error": f"Processing failed: {str(e)}"}), 500

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
