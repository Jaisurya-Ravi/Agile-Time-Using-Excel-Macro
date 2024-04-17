# Automated Agile Timesheet Management

Welcome to the Agile Timesheet Automation repository! Here, you will find the necessary steps to automate your Agile Timesheet process seamlessly.
</br>
</br>

# 1 Prerequisites
Ensure you have the following prerequisites installed on your system:

1. Microsoft Office 2007 or newer.
2. Selenium Basic: [Download](https://github.com/florentbr/SeleniumBasic/releases/download/v2.0.9.0/SeleniumBasic-2.0.9.0.exe)
3. EA Sendmail Component Manager: [Download](https://www.emailarchitect.net/webapp/download/easendmail.exe)
4. Browser Driver (e.g., Chromedriver): [Download](https://chromedriver.chromium.org/downloads)
</br>
</br>
</br>

# 2 steps to configure permissions

1. Open Excel by searching for it in the start menu.
2. Navigate to File > Excel Options > Trust Center > Trust Center Settings > ActiveX Settings. Ensure that the options 'Enable all controls without restrictions and without prompting' and 'Safe mode' are enabled as shown in the screenshot below.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/14251564-4c2c-4682-b90e-941823320dd6)

3. Click 'OK' to save the changes.
4. Next, go to File > Excel Options > Trust Center > Trust Center Settings > Macro Settings. Make sure that 'Enable All macros' and 'Trust Access to the VBA project object model' are enabled, as shown in the screenshot.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/cf87a7ab-b991-425a-83cf-404d323dbff9)

5.  Click 'OK' again to save the changes.
</br>
</br>
</br>

# 3 Installation Steps

### *3.1 Selenium Basic*
1. Download the Selenium Basic executable from the provided link.
2. Install Selenium Basic.

### *3.2 EA Sendmail Component Manager*
1. Download the EA Sendmail Component Manager executable from the provided link.
2. Install EA Sendmail Component Manager.

### *2.3 Updation of chromedriver*
1. Launch Chrome browser or install it if not already installed.
2. Navigate to Settings > About Chrome and note the version number.
3. Download the compatible version of Chromedriver using the provided link based on your Chrome version.
4. Replace the existing Chromedriver in the installation path of Selenium Basic with the downloaded one.

   ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/bd2d4878-bfd6-4da6-a3eb-e2df1992f48b)
</br>
</br>
</br>

# 3 Clone this repository

Clone this repository to your local machine to access the macro-enabled Excel file or download it from releases

1. Make a new folder and open the command prompt.
2. Input the below command.

```bash
git clone https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro.git
```
Click Here to download ⇊ [Download](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/releases/download/AgileTime1/AgileTime.xlsm)

3. After completion, you will be able to access the macro-enabled Excel file.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/2ab0b2ed-b133-4609-b52f-0c219861c77b)


### *3.1 To grant security permission to the downloaded Excel sheet*

Right-click on the downloaded Excel file and select 'Properties.'
Ensure that the 'Unblock' option is checked under the general tab. Click 'Apply' and then 'OK' to enable security permissions for the Excel sheet, as shown in the screenshot below.
</br>
</br>
</br>

# 3 Enabling permissions and Configuration

## User Details - Enter your Agile Time username, password, email ID, and manager ID in the respective cells.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/f5e9ed4f-d0e1-4b72-95fc-c00f29d46bcd)

## Master Details - This sheet automatically updates eligible task details based on your Agile Time records when you click the "Update master data" button. ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/d0477370-642a-440a-9ad5-1fbb95e0eee2)


![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/57108827-6def-41b2-bb21-a0ecdc0b7c02)

## Timesheet Details - Fill out your task details in this sheet, and the system will automatically update your timesheet in your Agile Time account when you click the "Upload timesheet" button. ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/7bc09cda-91d8-48ae-bd6c-10cd38686889)


![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/810a1f33-c26d-478c-8f2b-89b7384858b2)

## Configure email settings such as SMTP address, host, from email ID, subject, and body. Clicking the "Mail my Timesheet" button will automatically send an email to your specified email ID. ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/2d3769d2-e65f-4d7c-85d8-266b45de4efa)


![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/8d8948d5-4f5b-411f-80db-977b692df1f4)


# Enabling of buttons

1. Start by opening the Excel spreadsheet that you downloaded, then navigate to File > Options.
2. Go to the Quick Access Toolbar option in the left pane.
3. Activate the "Show Quick Access Toolbar below the Ribbon" setting.
4. Click on OK to save your changes and finish.
5. You will now be able to see the button on your Excel spreadsheet. ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/ee347425-a74d-4303-b074-d84c2db48526)


   ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/1185318d-d5ed-4a24-bd3d-8a0f9d13bf19)




## Thank you for using Agile Timesheet Automation!

** For Queries jaisurya.r@agile-labs.com **
