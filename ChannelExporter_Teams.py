#Importing the libraries.
import device_chromedriver_autoinstaller # This package avoids the need to manually pass a Chromedriver
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException
from selenium.webdriver.chrome.options import Options
# from webdriver_manager.chrome import ChromeDriverManager

import time
import psutil
from urllib.error import URLError, HTTPError
from http.client import IncompleteRead

# Import Libraries.
import pandas as pd
from bs4 import BeautifulSoup
import os
from datetime import datetime, timedelta

#__________________________________________________

# 1) Create frunctions for tasks.
## 1.1) Auto-install correct ChromeDriver version for ease of use.

# Function to attempt chromedriver installation with retries
def install_chromedriver(retries=3, delay=5):
    """
    Installs Chromedriver with retry attempts.

    Parameters:
    retries (int): Number of retry attempts (default: 3).
    delay (int): Delay between attempts in seconds (default: 5).

    Raises:
    RuntimeError: If installation fails after all retries.
    """

    for attempt in range(retries):
        try:
            device_chromedriver_autoinstaller.install_chromedriver()  # Check and install chromedriver
            print("Chromedriver installed successfully.")
            return
        except (URLError, HTTPError, IncompleteRead) as e:
            print(f"Attempt {attempt + 1} failed: {e}. Retrying in {delay} seconds...")
            time.sleep(delay)
    raise RuntimeError("Failed to install chromedriver after multiple attempts.")




#__________________________________________________

## 1.2) Open browser and navigate to url.

# [FINAL]
# Functions to load existing profile into the chrome browser and access MS Teams.
# Function to check if the browser is already open
def is_browser_open(process_name='chrome'):
    """
    Check if the specified browser process is currently running.

    Parameters:
    process_name (str): The name of the browser process to check (default: 'chrome').

    Returns:
    bool: True if the browser is open, False otherwise.
    """
        
    for proc in psutil.process_iter(['pid', 'name']):
        if process_name in proc.info['name'].lower():
            return True
    return False


# Function to open the browser if not open and navigate to a URL with a saved profile
def open_browser_with_profile_and_navigate(url, user_data_dir, profile_directory, isHeadless, scalefactor, window_height, window_width, headless_scalefactor, headless_height, headless_width):
    """
    Open the browser with a specified user profile and navigate to a URL.

    Parameters:
    url (str): The URL to navigate to.
    user_data_dir (str): The directory where user data is stored.
    profile_directory (str): The specific profile directory to load.
    isHeadless (bool): True for headless mode, False for UI mode.
    scalefactor (float): Scale factor for the UI mode.
    window_height (int): Height of the browser window in UI mode.
    window_width (int): Width of the browser window in UI mode.
    headless_scalefactor (float): Scale factor for headless mode.
    headless_height (int): Height of the browser window in headless mode.
    headless_width (int): Width of the browser window in headless mode.

    Returns:
    WebDriver: The Chrome WebDriver instance.
    """

    options = webdriver.ChromeOptions()
    options.add_argument(f'user-data-dir={user_data_dir}')
    options.add_argument(f'--profile-directory={profile_directory}')
    options.add_argument("--high-dpi-support=1")

    if isHeadless:

        print("WARNING 1: If the new chats are not exported, please use the Browser UI and click on the sign in button on Teams UI banner.")
        print("\nWARNING 2: The TEAMS CHANNEL navigation button should be pinned to the sidebar, pin it - if it is unpinned.")
        print("\nHeadless browser mode initiated - NO Browser UI will be displayed.")
        print("To view the browser UI, pass the arguments (isHeadless = False).\n")
        
        options.add_argument("--headless")
        # options.add_argument("--force-device-scale-factor=0.3")
        # options.add_argument("--window-size=2400,19900")
        options.add_argument(f"--force-device-scale-factor={headless_scalefactor}")
        options.add_argument(f"--window-size={headless_width},{headless_height}")  
    
    else:
        print("WARNING 1: If the new chats are not exported, please use the Browser UI and click on the sign in button on Teams UI banner.")
        print("\nWARNING 2: The TEAMS CHANNEL navigation button should be pinned to the sidebar, pin it - if it is unpinned.")
        print("\nUI browser mode initiated - Browser UI will be displayed.")
        print("To use a Headless browser, pass the arguments (isHeadless = True).\n")
        print("\nPlease ensure ZOOM Level in UI = 70 percent or less.\n")

        options.add_argument(f"--window-size={window_width},{window_height}")
        # options.add_argument("--window-size=2000,20000")  
        options.add_argument(f"--force-device-scale-factor={scalefactor}")

    
    if is_browser_open():
        print("Browser is already open. Skipping new browser launch.")
        # driver = webdriver.Chrome(service=ChromeService(driver_path), options=options)

        # Auto fetch and use ChromeDriver.
        # Attempt to install chromedriver with retry logic
        install_chromedriver()
        driver = webdriver.Chrome(options=options)

        driver.get(url)
        print(f"Navigated to {url}")

    else:
        #options.add_argument("--window-size=1280,800")
        # driver = webdriver.Chrome(service=ChromeService(driver_path), options=options)
        
        # Auto fetch and use ChromeDriver.
        # Attempt to install chromedriver with retry logic
        install_chromedriver()
        driver = webdriver.Chrome(options=options)

        driver.get(url)
        print(f"Navigated to {url}")
    
    return driver




#__________________________________________________

## 1.3) reusable functions.

# Dynamic wait to be used in all functions.
def wait_for_element_to_be_visible(driver, xpath, max_retries=3, wait_time=10):

    """
    Wait for an element to become visible on the page.

    This function attempts to locate an element specified by the 
    XPath and waits for it to be visible, retrying a specified 
    number of times if not found.

    Parameters:
    driver (WebDriver): The Selenium WebDriver instance.
    xpath (str): The XPath of the element to wait for.
    max_retries (int): Maximum number of retry attempts (default: 3).
    wait_time (int): Maximum wait time for each attempt in seconds (default: 10).

    Returns:
    WebElement: The visible element if found.

    Raises:
    TimeoutException: If the element is not visible after the maximum retries.
    """
        
    retries = 0
    while retries < max_retries:
        try:
            element = WebDriverWait(driver, wait_time).until(
                EC.visibility_of_element_located((By.XPATH, xpath))
            )
            return element
        except TimeoutException:
            retries += 1
            if retries == max_retries:
                raise
            print(f"Retrying... ({retries}/{max_retries})")





# Click on link.
def click_link(driver, xpath):
    """
    Click on a link identified by the specified XPath.

    This function waits for the link to become visible and attempts
    to click it. If the element is not found within the maximum 
    retry attempts, it handles the exception and prints an error message.

    Parameters:
    driver (WebDriver): The Selenium WebDriver instance.
    xpath (str): The XPath of the link to be clicked.

    Returns:
    WebDriver: The Selenium WebDriver instance after attempting to click the link.

    Raises:
    TimeoutException: If the element is not visible after the maximum retries.
    """

    # Implementing try and catch.
    try:
        link = wait_for_element_to_be_visible(driver, xpath, max_retries=5, wait_time=10)
        # link = driver.find_element(By.XPATH, xpath)
        link.click()
        return driver
    
    except TimeoutException:
        print("Element was not found after the maximum retries xpath - ", xpath)
    



# Enter text.
def enter_text(driver, xpath, input_text):
    """
    Enter text into a text field identified by the specified XPath.

    This function waits for the text field to become visible, clears 
    it, and then enters the provided text. If the element is not 
    found within the maximum retry attempts, it handles the exception 
    and prints an error message.

    Parameters:
    driver (WebDriver): The Selenium WebDriver instance.
    xpath (str): The XPath of the text field where the text will be entered.
    input_text (str): The text to be entered into the text field.

    Returns:
    WebDriver: The Selenium WebDriver instance after entering the text.

    Raises:
    TimeoutException: If the element is not visible after the maximum retries.
    """

    # Implementing try and catch.
    try:
        text_field = wait_for_element_to_be_visible(driver, xpath, max_retries=5, wait_time=10)
        # text_field = driver.find_element(By.XPATH, xpath)
        text_field.clear()
        text_field.send_keys(input_text)
        return driver
    
    except TimeoutException:
        print("Element was not found after the maximum retries xpath - ", xpath)
        


# Click button.
def click_button(driver, xpath):
    """
    Click on a button identified by the specified XPath.

    This function waits for the button to become visible and attempts 
    to click it. If the element is not found within the maximum retry 
    attempts, it handles the exception and prints an error message.

    Parameters:
    driver (WebDriver): The Selenium WebDriver instance.
    xpath (str): The XPath of the button to be clicked.

    Returns:
    WebDriver: The Selenium WebDriver instance after attempting to click the button.

    Raises:
    TimeoutException: If the element is not visible after the maximum retries.
    """

    # Implementing try and catch.
    try:
        button = wait_for_element_to_be_visible(driver, xpath, max_retries=5, wait_time=10)
        # button = driver.find_element(By.XPATH, xpath)
        button.click()
        return driver
    
    except TimeoutException:
        print("Element was not found after the maximum retries xpath - ", xpath)
    


# Scroll up and save html. - OLD
def scroll_to_top_and_save_html(driver):
    """
    Scroll to the top of the page and save the HTML content.

    This function continuously scrolls up the page, collecting the HTML 
    source at each step, until no new content is loaded. It returns 
    the accumulated HTML as a string.

    Parameters:
    driver (WebDriver): The Selenium WebDriver instance.

    Returns:
    str: The accumulated HTML content of the page after scrolling to the top.
    """

    htmlstring = ""
    while True:
        html = driver.page_source
        htmlstring += html + "\n"
        
        body = driver.find_element(By.TAG_NAME, 'body')
        body.send_keys(Keys.PAGE_UP)
        
        time.sleep(0.6)  # Wait for the scroll action to complete
        
        new_html = driver.page_source
        if html == new_html:
            break
    
    return htmlstring



#__________________________________________________

# 2) Functions to process the html and extract the feedbacks as a DF.
## 2.1) Function to Read html from text file.

# Read html from text file.
# Function to read HTML content from a text file
# Function to read HTML content from a text file
def read_html_content(file_path):
    # Check if the file exists
    if not os.path.exists(file_path):
        print(f"Error: The file '{file_path}' does not exist.")
        return None  # Return None if the file does not exist

    with open(file_path, 'r', encoding='utf-8') as file:
        # Display the imported HTML string
        print("The html file exists. File name - ", file_path)
        html_content = file.read()  # Read the entire file content as a string
    return html_content


#---------------------------

# Check if the DataFrame 'df' exists; if not, create it
def createDF(df_name):

    # df_name is the name of the DF.

    if df_name not in globals():

        print("\nCreating an empty DF because the provided df_name: '" + str(df_name) + "' - does NOT exist as a DF.\n\n")
        globals()[df_name] = pd.DataFrame()
    
    else:
        
        print("\nThe DF with the name: '" + str(df_name) + "' - already exists.\n\n")




#__________________________________________________
## 2.2) Function to Save the feedbacks to an excel


# Final one.

# If Excel output is necessary.
def save_to_excel(df_name):
    
    # Fetch directory where script is located.
    #script_path = os.path.dirname(os.path.abspath(__file__))
    #script_path = os.getcwd()
    
    # USe the df_name as filename.
    # excel_path = os.path.join(script_path, df_name + ".xlsx")
    excel_path = df_name

    globals()[df_name].to_excel(excel_path, index = False)
    print("The excel file has been saved.\nFile path:   ", excel_path)




# import pandas as pd
# from datetime import datetime, timedelta, time

def process_datetime_column(df, datetime_column):
    # Current date and time for reference
    current_datetime = datetime.now()
    today_str = current_datetime.strftime('%B %d, %Y')
    yesterday_str = (current_datetime - timedelta(days=1)).strftime('%B %d, %Y')

    # Function to convert the DateTime string
    def convert_datetime(dt_str):
        # Ensure dt_str is a string before processing
        if isinstance(dt_str, pd.Timestamp):
            dt_str = dt_str.strftime('%B %d, %Y %I:%M %p')  # Convert Timestamp to string if needed
        
        if isinstance(dt_str, str):  # Proceed if dt_str is a string
            if 'Yesterday' in dt_str:
                time_part = dt_str.split('at')[1].strip()
                return f'{yesterday_str} {time_part}'
            elif 'Today' in dt_str:
                time_part = dt_str.split('at')[1].strip() if 'at' in dt_str else dt_str.strip()
                return f'{today_str} {time_part}'
        
        return dt_str  # Return the original value if it's not a string

    # Apply the conversion function to the DateTime column
    df['DateTime'] = df[datetime_column].apply(convert_datetime)

    # Parse different formats including '%B %d, %Y %I:%M %p' and '%I:%M %p'
    def parse_datetime(dt_str):
        if isinstance(dt_str, str):  # Only attempt to parse if it's a string
            try:
                return datetime.strptime(dt_str, '%B %d, %Y %I:%M %p')
            except ValueError:
                try:
                    return datetime.strptime(dt_str, '%I:%M %p')
                except ValueError:
                    return None  # Return None if parsing fails
        return None  # Return None if not a string

    # Convert the DateTime column to datetime type
    df['DateTime'] = df['DateTime'].apply(parse_datetime)

    # Function to update dates with the year 1900 to today's date
    def update_date(dt):
        if isinstance(dt, datetime) and dt.year == 1900:
            return datetime.combine(current_datetime.date(), dt.time())
        return dt

    # Apply the update function to the DateTime column
    df['DateTime'] = df['DateTime'].apply(update_date)

    # Create new DATE and TIME columns
    df['DATE'] = df['DateTime'].dt.date
    df['TIME'] = df['DateTime'].dt.time

    return df


#---------------------------


def Feedback_df(output_filename, htmlstring, isOutputExcel):

    # df_name = Name of the DF.
    # html_filepath = Your HTML filepath as a string

    # Read the HTML content
    # html_content = read_html_content(html_filepath)
    html_content = htmlstring
    df_name = output_filename
    
    # Parse the HTML content
    soup = BeautifulSoup(html_content, 'html.parser')

    # Prepare lists to hold the extracted data
    data = []

    # Find all message containers using a more general search
    messages = soup.find_all('div', class_='fui-Flex')

    # Create or use an existing DataFrame
    createDF(df_name)

    print('Checking count - ', globals()[df_name].count())

    for message in messages:
        # Extract the author's name
        author = message.find('span', class_=lambda x: x and 'StyledText' in x)
        author_name = author.get_text(strip=True) if author else None

        # Extract the subject line
        subject_line = message.find('h2', {'data-tid': 'subject-line'})
        subject = subject_line.get_text(strip=True) if subject_line else ''
        
        # Extract the message content and preserve line breaks
        content_div = message.find('div', id=lambda x: x and x.startswith('content-'))
        if content_div:
            # Check if there are any <p> tags
            paragraphs = content_div.find_all('p')
            if paragraphs:
                # Join the text of each <p> tag with newline characters to preserve line breaks
                # content = '\n\n'.join([p.get_text(strip=True) for p in paragraphs])
                content = '\n'.join([p.get_text(strip=True) for p in paragraphs])
            else:
                # If no <p> tags, just get the text content of the div
                content = content_div.get_text(strip=True)
        else:
            content = None

        
        # Extract the datetime from the aria-label attribute
        timestamp = message.find('time', {'data-tid': 'timestamp'})
        datetime = timestamp['aria-label'] if timestamp and 'aria-label' in timestamp.attrs else None
        
        # Append extracted data to the list
        if author_name and content and datetime:
            data.append({
                'Name': author_name,
                'Subject': subject,
                'Content': content,
                'DateTime': datetime
            })

            # Create a new row as a DataFrame
            new_row = pd.DataFrame({
                'Name': [author_name],
                'Subject': [subject],
                'Content': [content],
                'DateTime': [datetime]
            })

            # Append the new row to the DataFrame
            globals()[df_name] = pd.concat([globals()[df_name] , new_row], ignore_index=True)


            # Append the new row to the DataFrame
            globals()[df_name] = pd.concat([globals()[df_name], new_row], ignore_index=True)

    # # Remove all duplicates (only unique rows)
    # globals()[df_name] = globals()[df_name].drop_duplicates().reset_index(drop=True)

    # Process the DateTime column
    globals()[df_name] = process_datetime_column(globals()[df_name], 'DateTime')

    # Regular expression to extract the correlation ID
    correlation_id_pattern = r'Correlation ID:\s*\[([0-9a-fA-F-]+)\]|Correlation Id:\s*([0-9a-fA-F-]+)|CorrelationId:\s*([0-9a-fA-F-]+)|ID:\s*([0-9a-fA-F-]+)|CorrelationID:\s*([0-9a-fA-F-]+)|\b([0-9a-fA-F-]{8}-[0-9a-fA-F-]{4}-[0-9a-fA-F-]{4}-[0-9a-fA-F-]{4}-[0-9a-fA-F-]{12})\b'

    # Use str.extract to extract the correlation ID and create a new column
    extracted_ids = globals()[df_name]['Content'].str.extract(correlation_id_pattern)

    # Combine all the extracted groups into a single column, replacing NaN with empty strings
    globals()[df_name]['Correlation ID'] = extracted_ids.bfill(axis=1).iloc[:, 0]

    # Remove all duplicates (only unique rows)
    globals()[df_name] = globals()[df_name].drop_duplicates().reset_index(drop=True)
    
    print("\n\nThe count of the DF = \n\n", globals()[df_name].count())

    # If excel output is necessary, save an excel.
    if isOutputExcel == True:
        save_to_excel(df_name)
    
    else:
        return globals()[df_name]



#__________________________________________________

# 3) Open browser and Navigate in Teams.
## 3.1) NS Teams -> Channel -> Sub Channel -> Excel

# Information to the user.
def information_chrome_userdirectory():

    print("\n------------------\nImportant Information:\n------------------\n")
    print("Prerequisites: Chrome installed and configured on the machine.")
    print("To identify the user data directory and profile directory in your device's Google Chrome using Chrome flags:")
    print("1. Open Chrome and log into the url (Teams)")
    print("2. In the address bar, type 'chrome://flags' and press Enter.")
    print("3. Use the search bar on the Chrome flags page to look for 'user data directory' or other related settings.")
    print("4. Replace 'C:/Path/To/Your/User Data' with your actual user data directory path, and 'ProfileName' with your profile directory.")



#__________________________________________________
# Function to handle the Sign in.
def click_sign_in(driver, signin_xpath):

    try:
        # Click on the sign in button automatically if needed:
        button = wait_for_element_to_be_visible(driver, xpath=signin_xpath, max_retries=2, wait_time=2)
        
        # Check if button is not None before clicking
        if button is not None:
            button.click()
            print("Sign in button found and clicked.\n")
        else:
            print("Sign in button not found.\n")
    except Exception as e:
        # Handle any exceptions that may occur during the process
        print("Sign in button not found.\n")

    return driver

#__________________________________________________

# Function for one click execution -  Output Excel.
def Chrome_export_TeamsChannel_excel(user_data_dir, profile_directory, output_filename, channelName, subChannelName, url= 'https://teams.microsoft.com', isHeadless = True, teams_button_xpath = "//button[@aria-label='Teams']", channel_pane_xpath = "//div[@data-tid='app-layout-area--main']", signin_xpath = "//div[contains(@class, 'fui-MessageBarActions')]//button[text()='Sign in']", scalefactor = "0.70", window_height="1080", window_width="1920", headless_scalefactor = "0.3", headless_height = "19900", headless_width = "2400"):
    """
    Export data from a Microsoft Teams channel to an Excel file.

    This function launches a Chrome browser with the specified user profile,
    navigates to Microsoft Teams, expands the specified channel and sub-channel,
    scrolls to the top of the page, and saves the HTML content. The output is 
    processed into an Excel file.

    Parameters:
    user_data_dir (str): The directory for Chrome user data.
    profile_directory (str): The profile directory to use in Chrome.
    output_filename (str): The filename for the output Excel file.
    channelName (str): The name of the Teams channel to navigate to.
    subChannelName (str): The name of the sub-channel to navigate to.
    url (str): The URL to navigate to (default: 'https://teams.microsoft.com').
    isHeadless (bool): If True, run the browser in headless mode (default: True).
    teams_button_xpath (str): XPath to the Teams button.
    channel_pane_xpath (str): XPath to the channel pane.
    signin_xpath = XPath to the Sign in button sometimes visible.
    scalefactor (str): Scale factor for the UI (default: "0.70").
    window_height (str): Height of the browser window (default: "1080").
    window_width (str): Width of the browser window (default: "1920").
    headless_scalefactor (str): Scale factor for headless mode (default: "0.3").
    headless_height (str): Height of the browser window in headless mode (default: "19900").
    headless_width (str): Width of the browser window in headless mode (default: "2400").

    Returns:
    Excel: The output in Excel format.
    """

    information_chrome_userdirectory()    #Prints the important information to Obtain the Chrome user directory..
    print('\n------------------\nLogs:\n------------------\n')

    # Launch the browser with the user profile.
    driver = open_browser_with_profile_and_navigate(url, user_data_dir, profile_directory, isHeadless = isHeadless, scalefactor = scalefactor, window_height=window_height, window_width=window_width, headless_scalefactor = headless_scalefactor, headless_height = headless_height, headless_width = headless_width)

    # Click on the 'Teams' button to navigate to the 'Teams' chat section.
    driver = click_button(driver, xpath=teams_button_xpath)

    # Click on the sign in button automatically if needed:
    driver = click_sign_in(driver, signin_xpath)

    # Use WebDriverWait to wait for the element to be present
    element_xpath = f"//div[@displayname='{channelName}']"
    channel_xpath = f"//div[@data-tid='channel-list-team-node']//span[contains(text(),'{channelName}') and @dir='auto']"
    element = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, element_xpath)))
    
    # Extract the id attribute
    id_attribute = element.get_attribute('id')
    is_expanded = element.get_attribute('aria-expanded')

    if is_expanded == 'false':
        driver = click_button(driver, xpath=channel_xpath)
        print(f'\n\n 1) Clicked the channel and expanded the dropdown. - {channelName}')
        print(f"\n    1.1) The extracted id attribute is: {id_attribute}")

    elif is_expanded == 'true':
        print(f' 1) Channel is already clicked and expanded. - {channelName}')

    else:
        print(f'\n\n ERROR - Could not locate the channel name = {channelName}')
        print(f'\n The xpath used is = {channel_xpath}')
        print(f'\n The id_attribute = {id_attribute}')
        print(f'\n The is_expanded tag = {is_expanded}')

    # Use the extracted id and Click on the subchannel.    
    subChannel_xpath = f"//div[contains(@id, '{id_attribute}')]//span[contains(text(),'{subChannelName}')]"
    driver = click_button(driver, xpath=subChannel_xpath)
    print(f'\n\n 2) Clicked the sub channel - {subChannelName}')
    
    # Focus on the channel pane before scrolling
    driver.execute_script("arguments[0].focus();", WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, channel_pane_xpath))))
    driver = click_button(driver, xpath = channel_pane_xpath)

    # Scroll to the end of the page
    # Scroll to the top of the page and save HTML
    htmlstring = scroll_to_top_and_save_html(driver)

    time.sleep(2)
    driver.quit()

    print('\n------------\nBrowser closed.\n------------\n')
    print('\n DF or Excel Feedback consolidation in progress - This can take up to or more than 10 minutes.')
    print('Once it is completed, further details will be furnished below.\n\n')

    print('Output format = Excel.')
    return Feedback_df(output_filename, htmlstring, isOutputExcel = True)



#__________________________________________________
## 3.2) NS Teams -> Channel -> Sub Channel -> DF


# Function for one click execution -  Output = DF.
def Chrome_export_TeamsChannel_df(user_data_dir, profile_directory, output_filename, channelName, subChannelName, url= 'https://teams.microsoft.com', isHeadless = True, teams_button_xpath = "//button[@aria-label='Teams']", signin_xpath = "//div[contains(@class, 'fui-MessageBarActions')]//button[text()='Sign in']", channel_pane_xpath = "//div[@data-tid='app-layout-area--main']", window_height="1080", window_width="1920", scalefactor = "0.70", headless_scalefactor = "0.3", headless_height = "19900", headless_width = "2400"):
    """Export data from a Microsoft Teams channel to an Excel file.

    This function launches a Chrome browser with the specified user profile,
    navigates to Microsoft Teams, expands the specified channel and sub-channel,
    scrolls to the top of the page, and saves the HTML content. The output is 
    processed into an Excel file.

    Parameters:
    user_data_dir (str): The directory for Chrome user data.
    profile_directory (str): The profile directory to use in Chrome.
    output_filename (str): The filename for the output Excel file.
    channelName (str): The name of the Teams channel to navigate to.
    subChannelName (str): The name of the sub-channel to navigate to.
    url (str): The URL to navigate to (default: 'https://teams.microsoft.com').
    isHeadless (bool): If True, run the browser in headless mode (default: True).
    teams_button_xpath (str): XPath to the Teams button.
    channel_pane_xpath (str): XPath to the channel pane.    
    signin_xpath = XPath to the Sign in button sometimes visible.
    scalefactor (str): Scale factor for the UI (default: "0.70").
    window_height (str): Height of the browser window (default: "1080").
    window_width (str): Width of the browser window (default: "1920").
    headless_scalefactor (str): Scale factor for headless mode (default: "0.3").
    headless_height (str): Height of the browser window in headless mode (default: "19900").
    headless_width (str): Width of the browser window in headless mode (default: "2400").

    Returns:
    DataFrame: The output is a DataFrame."""

    information_chrome_userdirectory()    #Prints the important information to Obtain the Chrome user directory..
    print('\n------------------\nLogs:\n------------------\n')

    # Launch the browser with the user profile.
    driver = open_browser_with_profile_and_navigate(url, user_data_dir, profile_directory, isHeadless = isHeadless, scalefactor = scalefactor, window_height=window_height, window_width=window_width, headless_scalefactor = headless_scalefactor, headless_height = headless_height, headless_width = headless_width)

    # Click on the 'Teams' button to navigate to the 'Teams' chat section.
    driver = click_button(driver, xpath=teams_button_xpath)
    # Click on the sign in button automatically if needed:

    driver = click_sign_in(driver, signin_xpath)
    # Use WebDriverWait to wait for the element to be present

    element_xpath = f"//div[@displayname='{channelName}']"
    channel_xpath = f"//div[@data-tid='channel-list-team-node']//span[contains(text(),'{channelName}') and @dir='auto']"
    element = WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, element_xpath)))
    
    # Extract the id attribute
    id_attribute = element.get_attribute('id')
    is_expanded = element.get_attribute('aria-expanded')

    if is_expanded == 'false':
        driver = click_button(driver, xpath=channel_xpath)
        print(f'\n\n 1) Clicked the channel and expanded the dropdown. - {channelName}')
        print(f"\n    1.1) The extracted id attribute is: {id_attribute}")

    elif is_expanded == 'true':
        print(f' 1) Channel is already clicked and expanded. - {channelName}')

    else:
        print(f'\n\n ERROR - Could not locate the channel name = {channelName}')
        print(f'\n The xpath used is = {channel_xpath}')
        print(f'\n The id_attribute = {id_attribute}')
        print(f'\n The is_expanded tag = {is_expanded}')

    # Use the extracted id and Click on the subchannel.    
    subChannel_xpath = f"//div[contains(@id, '{id_attribute}')]//span[contains(text(),'{subChannelName}')]"
    driver = click_button(driver, xpath=subChannel_xpath)
    print(f'\n\n 2) Clicked the sub channel - {subChannelName}')
    
    # Focus on the channel pane before scrolling
    driver.execute_script("arguments[0].focus();", WebDriverWait(driver, 7).until(EC.presence_of_element_located((By.XPATH, channel_pane_xpath))))
    driver = click_button(driver, xpath = channel_pane_xpath)

    # Scroll to the end of the page
    # Scroll to the top of the page and save HTML
    htmlstring = scroll_to_top_and_save_html(driver)

    time.sleep(2)
    driver.quit()

    print('\n------------\nBrowser closed.\n------------\n')
    print('\n DF or Excel Feedback consolidation in progress - This can take up to or more than 10 minutes.')
    print('Once it is completed, further details will be furnished below.\n\n')

    # Return the df if a DF is requested.
    print('Output format = Dataframe.')
    return Feedback_df(output_filename, htmlstring, isOutputExcel = False)





#__________________________________________________
