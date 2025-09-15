# Web Scraper for Editor Comments

This project uses "Selenium" to log into the [MMAI ATU system](https://mmai.atu.ac.ir/), navigate through the "Editor (دبیر علمی)" role, and extract the "latest comments" attached to accepted papers.  
The results are exported to an Excel file with the title "reviewer_final_comments.xlsx" containing both the paper codes and the corresponding editor comments.

## Requirements

- **Python** 3.8 or higher
- **Google Chrome** (latest stable release recommended)
- **ChromeDriver** (must match your installed Chrome version)

 To check your Chrome version:
- On Windows: open chrome://settings/help
- If needed download the matching ChromeDriver from:[https://chromedriver.chromium.org/downloads](https://chromedriver.chromium.org/downloads)

## Python Dependencies

To install the required Python libraries, follow these steps:

Open your command prompt (CMD) or terminal.

It is recommended to disable any VPN or proxy connections to avoid installation issues.

Run the following command:

pip install selenium pandas openpyxl

The download and installation will begin immediately.

If a library is already installed, you will see a message similar to:
Requirement already satisfied: selenium in ...
