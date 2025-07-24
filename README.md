# **BISE Sargodha Matric Result Scraper**

This Python script is designed to scrape Matriculation exam results from the BISE Sargodha official website for a given range of roll numbers. It retrieves student details and subject-wise marks, then compiles this data into a well-formatted Excel file, highlighting subjects where a student has failed.

## **Features**

* **Batch Result Retrieval:** Fetches results for a user-defined range of roll numbers.  
* **Detailed Data Extraction:** Extracts Roll Number, Candidate Name, Father Name, and individual subject marks.  
* **Organized Output:** Stores results in an Excel file (.xlsx) with a clear, predefined column order.  
* **Intelligent Highlighting:** Automatically identifies and highlights (in light red) the cells for subjects in which a student has failed, based on the "overall result" string.  
* **Enhanced Excel Formatting:** Applies bold headers, auto-adjusts column widths, and freezes the header row and first two columns for improved readability and navigation.

## **Prerequisites**

Before running this script, ensure you have Python 3 installed on your system.  
You will also need the following Python packages:

* requests: For making HTTP requests to the website.  
* beautifulsoup4: For parsing HTML content.  
* pandas: For data manipulation and Excel export.  
* openpyxl: Backend engine for pandas to read/write .xlsx files and for advanced Excel formatting.

## **Installation**

1. **Clone the repository (or download the results.py file):**  
   git clone https://github.com/your-username/your-repo-name.git  
   cd your-repo-name

   (Replace your-username and your-repo-name with your actual GitHub details.)  
2. Install the required Python packages:  
   Open your terminal or command prompt and run:  
   pip install requests beautifulsoup4 pandas openpyxl

## **How to Use**

1. Run the script:  
   Navigate to the directory where you saved results.py in your terminal or command prompt and execute:  
   python results.py

2. Enter Roll Number Range:  
   The script will prompt you to enter the starting and ending roll numbers:  
   Enter the starting roll number (e.g., 520001): 520001  
   Enter the ending roll number (e.g., 520010): 520010

   Enter the desired range and press Enter after each prompt.  
3. Result Retrieval:  
   The script will then proceed to retrieve results for each roll number in the specified range. You will see progress messages in the console.  
4. Excel Output:  
   Once the process is complete, an Excel file named bise\_matric\_results.xlsx will be created or updated in the same directory where you ran the script.

## **Excel Output Structure**

The generated Excel file will have the following columns in order:

* Roll-No  
* Candidate Name  
* Father Name  
* Computer Science  
* Biology  
* Mathematics  
* Physics  
* Chemistry  
* Islamiyat  
* Pakistan Studies  
* Urdu  
* English  
* THQ (Translation of Holy Quran)  
* overall result

**Highlighting:** Cells corresponding to subjects where a student has failed (indicated by abbreviations like BIO, PHY, CHM, EGL, URU, MAT, THQ, PKS, ISM, CSC, CS in the 'overall result' column, including variations like BIOI, BIOII, BIO(PR), ISMI, ISMII) will be highlighted in light red.  
**Example of Excel Output:**  
*(Please insert the screenshot you provided here, e.g., by uploading it to your GitHub repository and linking it like this: \!\[Example Excel Output\](path/to/your/image\_668821.png). Make sure the image is accessible in your repo.)*

## **Future Enhancements**

This script forms the foundation for a more robust application. Future plans include developing an Android application to provide a user-friendly interface for result retrieval, local SQLite storage, and in-app viewing and export capabilities.