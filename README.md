# Automated Agile Timesheet Management

Welcome to the Agile Timesheet Automation repository! Here, you will find the necessary steps to automate your Agile Timesheet process seamlessly.

## Prerequisites
Ensure you have the following prerequisites installed on your system:

- Microsoft Office 2007 or newer.
- Selenium Basic: [Download](https://github.com/florentbr/SeleniumBasic/releases/download/v2.0.9.0/SeleniumBasic-2.0.9.0.exe)).
- EA Sendmail Component Manager: [Download](https://www.emailarchitect.net/webapp/download/easendmail.exe)
- Browser Driver (e.g., Chromedriver): [Download](https://chromedriver.chromium.org/downloads)

# Installation Steps

## Selenium Basic
1. Download the Selenium Basic executable from the provided link.
2. Install Selenium Basic.

## EA Sendmail Component Manager
1. Download the EA Sendmail Component Manager executable from the provided link.
2. Install EA Sendmail Component Manager.

## Configuration of chromedriver
1. Launch Chrome browser or install it if not already installed.
2. Navigate to Settings > About Chrome and note the version number.
3. Download the compatible version of Chromedriver using the provided link based on your Chrome version.
4. Replace the existing Chromedriver in the installation path of Selenium Basic with the downloaded one.

## Clone this repository

Clone this repository to your local machine to access the macro-enabled Excel file:

```bash
https://github.com/jaisuryaravi/SeleniumAutomationUDY.git
```bash

# Screenshots
## User Details - Enter your Agile Time username, password, email ID, and manager ID in the respective cells.
## Master Details - This sheet automatically updates eligible task details based on your Agile Time records when you click the "Update master data" button.
## Timesheet Details - Fill out your task details in this sheet, and the system will automatically update your timesheet in your Agile Time account when you click the "Upload timesheet" button.
## Configure email settings such as SMTP address, host, from email ID, subject, and body. Clicking the "Mail my Timesheet" button will automatically send an email to your specified email ID.


Thank you for using Agile Timesheet Automation!
