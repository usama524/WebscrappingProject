from flask import Flask, render_template, request, send_file, jsonify, session
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, NoSuchElementException
import datetime
import os

app = Flask(__name__)
app.secret_key = os.urandom(24)  # Needed for session management

def fetch_driver_expiry_date(driver_licence_no, driver):
    try:
        full_driver_licence_no = str(driver_licence_no)  
        search_driver_licence_no = full_driver_licence_no  # Use full badge number directly

        print(f"Searching for driver with full badge number: {search_driver_licence_no}")

        # Navigate to the driver licence search page
        driver.get("https://tph.tfl.gov.uk/TfL/SearchDriverLicence.page?org.apache.shale.dialog.DIALOG_NAME=TPHDriverLicence&Param=lg2.TPHDriverLicence&menuId=6")
        
        # Wait for the input field to be available and then fill it
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchdriverlicenceform:DriverLicenceNo")))

        # Clear and input the full driver licence number
        driver_licence_input = driver.find_element(By.ID, "searchdriverlicenceform:DriverLicenceNo")
        driver_licence_input.clear()
        driver_licence_input.send_keys(search_driver_licence_no)

        # Wait for the submit button to be clickable and click it
        submit_button = WebDriverWait(driver, 20).until(
            EC.element_to_be_clickable((By.ID, "searchdriverlicenceform:_id189"))
        )
        submit_button.click()

        # Wait for the next page to load and check for results
        WebDriverWait(driver, 20).until(EC.url_contains("https://tph.tfl.gov.uk/TfL/lg2/TPHLicensing/pubregsearch/Driver/SearchDriverLicence.page"))

        # Wait for the driver results table to appear
        try:
            # This is the tbody containing the results
            WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "_id177:driverResults:tbody_element")))
        except TimeoutException:
            return "Error: Could not find the driver results."

        # Now that the table is present, let's extract the expiry date
        rows = driver.find_elements(By.CSS_SELECTOR, "#_id177\\:driverResults\\:tbody_element tr")

        for row in rows:
            columns = row.find_elements(By.TAG_NAME, "td")
            if len(columns) >= 3:  # If the row has at least 3 columns
                expiry_date = columns[2].text.strip()  # The expiry date is in the 3rd column
                return expiry_date

        return "Expiry Date Not Found"

    except TimeoutException as e:
        print(f"TimeoutException: The page took too long to load: {str(e)}")
        return "Error: Timeout"
    
    except NoSuchElementException as e:
        print(f"NoSuchElementException: Element not found: {str(e)}")
        return "Error: Element not found"
    
    except Exception as e:
        print(f"Unexpected error: {str(e)}")
        return "Revoked & suspended to work"

# Function to fetch expiry date for vehicles (using VRM)
def fetch_expiry_date(vrm, driver):
    try:
        driver.get("https://tph.tfl.gov.uk/TfL/SearchVehicleLicence.page?org.apache.shale.dialog.DIALOG_NAME=TPHVehicleLicence&Param=lg2.TPHVehicleLicence&menuId=7")
        WebDriverWait(driver, 20).until(EC.presence_of_element_located((By.ID, "searchvehiclelicenceform:VehicleVRM")))
        vrm_input = driver.find_element(By.ID, "searchvehiclelicenceform:VehicleVRM")
        vrm_input.clear()
        vrm_input.send_keys(vrm)
        submit_button = driver.find_element(By.ID, "searchvehiclelicenceform:_id187")
        submit_button.click()
        WebDriverWait(driver, 20).until(EC.url_contains("https://tph.tfl.gov.uk/TfL/lg2/TPHLicensing/pubregsearch/Vehicle/SearchVehicleLicence.page"))
        
        error_message = driver.find_elements(By.ID, "validation")
        if error_message:
            return "Revoked & suspended to work"

        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "_id173")))

        paragraphs = driver.find_elements(By.TAG_NAME, "p")
        for para in paragraphs:
            if "Licence Expiry Date:" in para.text:
                expiry_date = para.text.split("Licence Expiry Date:")[1].strip()
                return expiry_date
        return "Expiry Date Not Found"
    except Exception as e:
        return "Error fetching expiry date"
    

def delete_old_files(is_drivers):
    # Define directories for driver and vehicle
    driver_dir = 'downloadedFiles/driver'
    vehicle_dir = 'downloadedFiles/vehicle'

    # Delete files based on whether it's for drivers or vehicles
    if is_drivers:
        # Delete all files in the driver folder
        for filename in os.listdir(driver_dir):
            file_path = os.path.join(driver_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print("Deleted old files in 'downloadedFiles/driver'")
    else:
        # Delete all files in the vehicle folder
        for filename in os.listdir(vehicle_dir):
            file_path = os.path.join(vehicle_dir, filename)
            if os.path.isfile(file_path):
                os.remove(file_path)
        print("Deleted old files in 'downloadedFiles/vehicle'")

def process_xlsx(file, driver, is_drivers=False):
    # Delete old files based on the type (driver or vehicle)
    delete_old_files(is_drivers)

    df = pd.read_excel(file)
    
    # Check if the required columns are available
    if df.shape[1] <= 4:
        return "Error: Required columns (Column E) are not available in this file."
    
    total_rows = df.shape[0]
    session['total_rows'] = total_rows  # Set total rows for progress tracking
    session['processed_rows'] = 0  # Initialize processed rows to 0
    
    completed_rows = 0  # To track the rows processed so far
    
    # Iterate through each row in the DataFrame
    for index, row in df.iterrows():
        if is_drivers:
            badge_number = str(row.iloc[3]).strip()  # Column D - Badge Number
            if len(badge_number) <= 4:
                continue  # Skip rows where badge number length is invalid
            
            # Truncate the badge number (remove last 4 digits) for searching the expiry date
            search_badge_number = badge_number[:-4]
            
            # Fetch the expiry date using the truncated badge number
            expiry_date = fetch_driver_expiry_date(search_badge_number, driver)
            
            # Update the expiry date in Column E (or the correct column)
            df.at[index, df.columns[4]] = expiry_date  # Column E
        else:
            vrm = str(row.iloc[3]).replace(" ", "").strip()  # Column E - VRM
            if not vrm:
                continue  # Skip rows where VRM is missing or empty
            
            # Fetch the expiry date for the car based on VRM
            expiry_date = fetch_expiry_date(vrm, driver)
            
            # Update the expiry date in Column F (or the correct column for cars)
            df.at[index, df.columns[5]] = expiry_date  # Column F (Change this if necessary)
        
        # Update progress tracking
        completed_rows += 1
        session['processed_rows'] = completed_rows  # Update processed rows in session
        session.modified = True  # Ensure session is updated
        
    # Generate filename based on whether it's for drivers or vehicles
    current_time = datetime.datetime.now().strftime('%Y%m%d%H%M%S')
    if is_drivers:
        output_filename = f"DriverLicense-{current_time}.xlsx"
        output_path = os.path.join('downloadedFiles/driver', output_filename)
    else:
        output_filename = f"VehicleLicense-{current_time}.xlsx"
        output_path = os.path.join('downloadedFiles/vehicle', output_filename)
    
    # Save the DataFrame to Excel
    df.to_excel(output_path, index=False)
    
    return output_filename  # Return the path to the updated file


@app.route('/')
def index():
    is_drivers = request.args.get('is_drivers', 'false')  # Default to 'false' if not set
    return render_template('index.html', is_drivers=is_drivers)

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file uploaded"}), 400
    file = request.files['file']
    is_drivers = request.form.get('is_drivers', 'false') == 'true'  # Get the 'is_drivers' value from the form

    # Set up Selenium WebDriver (headless)
    options = webdriver.ChromeOptions()
    # options.add_argument("--headless")
    driver = webdriver.Chrome(options=options)

    # Process the file in a separate thread or background task
    output_filename = process_xlsx(file, driver, is_drivers)
    driver.quit()

    return jsonify({"download_url": f"/download/{output_filename}?is_drivers={str(is_drivers).lower()}"}), 200


@app.route('/progress')
def get_progress():
    # Send the current progress
    total = session.get('total_rows', 0)
    processed = session.get('processed_rows', 0)
    return jsonify({"processed": processed, "total": total})

@app.route('/download/<filename>')
def download_file(filename):
    # Get the 'is_drivers' query parameter (default to False if not provided)
    is_drivers = request.args.get('is_drivers', 'false') == 'true'
    
    # Determine the subdirectory based on 'is_drivers'
    if is_drivers:
        file_path = os.path.join('downloadedFiles', 'driver', filename)
    else:
        file_path = os.path.join('downloadedFiles', 'vehicle', filename)
    
    # Check if the file exists
    if not os.path.exists(file_path):
        return jsonify({"error": "File not found"}), 404
    
    # Send the file as a download response
    return send_file(file_path, as_attachment=True)

if __name__ == "__main__":
    app.run(debug=True)
