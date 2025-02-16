# ChannelExporter_Teams

## Overview

The **ChannelExporter_Teams** package allows users to automate interactions with Microsoft Teams, facilitating the export of channel feedback into structured formats like Excel or DataFrames. This package uses Selenium for browser automation and BeautifulSoup for HTML parsing.

## Features

- Seamless integration with Microsoft Teams.
- Export channel data to Excel files or DataFrames.
- Flexible parameters for browser configurations.
- Automatic management of Chromedriver installation.

## Installation

To install the ChannelExporter_Teams package, you can use pip:

```bash
pip install ChannelExporter_Teams
```

Ensure you have the required dependencies installed, which include `selenium`, `pandas`, `beautifulsoup4`, and `psutil`.

## Usage

### Function 1: Chrome_export_TeamsChannel_excel
This function exports data from a specified Microsoft Teams channel to an Excel file.

### Example 1: Using Default Arguments

You can call the function using the required parameters, allowing the use of default values for optional parameters.


```Python 
# Importing the library.
import ChannelExporter_Teams as cet

# Define required parameters
user_data_directory = r'C:/Users/YourUsername/AppData/Local/Google/Chrome/User Data'
profile_directory = 'Default'
output_file = 'teams_feedback.xlsx'
channel_name = 'Your Channel Name'
sub_channel_name = 'Your Sub Channel Name'

# Call the function with default arguments for optional parameters
cet.Chrome_export_TeamsChannel_excel(user_data_directory, profile_directory, output_file, channel_name, sub_channel_name)
```


### Example 2: Using Default and Advanced Keyword Arguments with Browser UI.

You can also customize the function call by specifying values for optional parameters.

```Python 
# Importing the library.
import ChannelExporter_Teams as cet

# Define your parameters
user_data_directory = r'C:/Users/YourUsername/AppData/Local/Google/Chrome/User Data'
profile_directory = 'Default'
output_file = 'teams_feedback_custom.xlsx'
channel_name = 'Your Channel Name'
sub_channel_name = 'Your Sub Channel Name'

# Call the function with custom settings
cet.Chrome_export_TeamsChannel_excel(
    user_data_dir=user_data_directory,
    profile_directory=profile_directory,
    output_filename=output_file,
    channelName=channel_name,
    subChannelName=sub_channel_name,
    isHeadless=False,  # Run the browser in UI mode
    window_height="1200",  # Customize window height
    window_width="1600"    # Customize window width
    scalefactor = "0.5",   # UI scaling factor
    teams_button_xpath = "//button[@aria-label='Teams']",    # futureproofing the xpaths
    channel_pane_xpath = "//div[@data-tid='app-layout-area--main']"    # futureproofing the xpaths.
)
```

### Function 2: Chrome_export_TeamsChannel_df
This function exports data from a specified Microsoft Teams channel to a DataFrame.


### Example 3: Headlessly and DataFrame - Using Default and Advanced Keyword Arguments with NO Browser UI.

To run the browser in headless mode (without a visible UI), simply set the isHeadless parameter to True. This is useful for automated scripts running in environments where a graphical interface may not be available.

The output is a Dataframe this time, the same settings can also be used with the excel method - 'Chrome_export_TeamsChannel_df'


```Python 
# Importing the library.
import ChannelExporter_Teams as cet

# Define your parameters for headless mode
user_data_directory = r'C:/Users/YourUsername/AppData/Local/Google/Chrome/User Data'
profile_directory = 'Default'
output_file = 'teams_feedback_headless.xlsx'
channel_name = 'Your Channel Name'
sub_channel_name = 'Your Sub Channel Name'

# Call the function with headless mode enabled
df_Output = cet.Chrome_export_TeamsChannel_df(
    user_data_dir=user_data_directory,
    profile_directory=profile_directory,
    output_filename=output_file,
    channelName=channel_name,
    subChannelName=sub_channel_name,
    isHeadless=True,  # Run the browser in headless mode
    headless_height="1080",  # Customize headless window height
    headless_width="1920",     # Customize headless window width
    headless_scalefactor = "0.5",  # Headless scale factor.
    teams_button_xpath = "//button[@aria-label='Teams']",    # futureproofing the xpaths
    channel_pane_xpath = "//div[@data-tid='app-layout-area--main']"    # futureproofing the xpaths.
)
```


