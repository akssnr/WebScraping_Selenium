                                             # # DOCUMENTATION


To ensure script runs smoothly, for setting up the environment and executing the script. Here’s a comprehensive guide to prepare your system for running the Selenium and DOCX processing script:

## Environment Setup Instructions:

### 1. Install Python
Make sure you have Python installed on your computer. You can download and install it from [python.org](https://www.python.org/downloads/). Ensure you add Python to your system's PATH during installation.

### 2. Install Required Packages
Open your command line interface (CLI) and install the necessary Python packages using pip. Execute the following commands:

### In Terminal:
pip install selenium
pip install python-docx
pip install pandas
pip install json

### 3. Download ChromeDriver
 Go to the [ChromeDriver download page](https://googlechromelabs.github.io/chrome-for-testing/) or https://sites.google.com/chromium.org/driver/.
 Choose the version of ChromeDriver that matches the version of Chrome installed on your system.
 Download the correct version for your operating system.
 Extract the downloaded ZIP file and note the path to the `chromedriver.exe` file. You will need to use this path in your script.

### 4. Set Up the Script
 Make sure the script file `data_extract_selenium.py` is saved on your local machine.
 Update the `executable_path` in the `setup_selenium_driver()` function to the path where you extracted `chromedriver.exe`.

### 5. Prepare the DOCX File
 Ensure that the DOCX file from which URLs will be extracted is located in a known directory.
 Update the ‘docx_file_path’ in the ‘main()’ function to the path of your DOCX file.
 Also, Update the ‘output_dir’ in the ‘main()’ function 


# Running the Script

Follow these steps to execute the script:

### 1. Open Command Line Interface (CLI):
    Navigate to the folder containing your script. You can use `cd path\to\your\script` in the command line to change to the correct directory.

### 2. Execute the Script:
    Run the script by typing the following command and then pressing Enter:
     # In Terminal:
     python extraction_selenium.py    

### 3. Check the Output:
    After the script completes execution, check the same directory for `output.json` and `output.xlsx` files containing the scraped data.

## Additional Tips
 Ensure your internet connection is active as the script needs to access websites.
•	If you encounter issues with ChromeDriver, verify that its version matches your Chrome browser's version and that it is correctly located in the specified path.
•	If you receive errors related to missing packages, make sure all required packages are installed as per the instructions above.
•	By following these setup and execution instructions, you should be able to run the script without issues and obtain the desired data in both JSON and Excel formats.
