import time
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from datetime import datetime
from selenium.webdriver.chrome.service import Service
from webdriver_manager.chrome import ChromeDriverManager 

# Install and Initialize the WebDriver
service = Service(ChromeDriverManager().install())  
driver = webdriver.Chrome(service=service)

# Here get google suggestions
def get_google_suggestions(query):
    driver.get("https://www.google.com/")
    search_box = driver.find_element(By.NAME, "q")
    search_box.send_keys(query)
    time.sleep(2)  # Wait for suggestions to load

    # Here get the suggestion elements
    suggestions = driver.find_elements(By.CSS_SELECTOR, "div.primary-side")
    
    longest = ""
    shortest = ""

    # Logic for longest and shortest suggestion
    for suggestion in suggestions:
        text = suggestion.text
        if len(text) > len(longest):
            longest = text
        if len(text) < len(shortest) or shortest == "":
            shortest = text

    return longest, shortest

def update_excel():
    # Setup the Excel file
    file_path = r'G:\Python\seleniumTable.xlsx'

    # Open the Excel file
    excel_data = pd.ExcelFile(file_path)

    # Get today's day of the week
    day_of_week = datetime.today().strftime('%A')

    # Logic for today exists otherwise create a new sheet
    if day_of_week not in excel_data.sheet_names:
        print(f"No data for today ({day_of_week}). Creating a new sheet.")
        
        df = pd.DataFrame(columns=["Keyword", "Longest", "Shortest"])
        with pd.ExcelWriter(file_path, engine="openpyxl", mode='a') as writer:
            df.to_excel(writer, sheet_name=day_of_week, index=False)
        print(f"New sheet for {day_of_week} created.")
        return

    # Load the sheet for today
    df = pd.read_excel(excel_data, sheet_name=day_of_week)

    # Loop through the keywords in the first column 
    for index, row in df.iterrows():
        keyword = row[0]

        print(f"Processing keyword: {keyword}")

        longest, shortest = get_google_suggestions(keyword)

        # Add the results to the columns for longest and shoetest value
        df.at[index, 'Longest'] = longest
        df.at[index, 'Shortest'] = shortest

    # Save the updated Excel file
    df.to_excel(file_path, sheet_name=day_of_week, index=False)

if __name__ == "__main__":
    try:
        update_excel()
    finally:
        driver.quit()  # Close the browser when done


# Here I use meaningfull variable name, function name and add comment for making meaningfull code.