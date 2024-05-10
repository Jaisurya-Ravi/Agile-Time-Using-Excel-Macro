# Automated Agile Timesheet Management

Welcome to the Agile Timesheet Automation repository! Here, you will find the necessary steps to automate your Agile Timesheet process seamlessly.
</br>
</br>

# 1 Prerequisites
Ensure you have the following prerequisites installed on your system by following the installation steps given in the below steps:

1. Microsoft Office 2007 or newer.
2. Selenium Basic
3. EA Sendmail Component Manager
4. Chrome Web Browser
5. Browser Driver (Chromedriver)
</br>
</br>
</br>

# 2 steps to enable permissions

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
1. Download the Selenium Basic [Download](https://github.com/florentbr/SeleniumBasic/releases/download/v2.0.9.0/SeleniumBasic-2.0.9.0.exe) executable from the provided link.
2. Install Selenium Basic.

### *3.2 EA Sendmail Component Manager*
1. Download the EA Sendmail Component Manager [Download](https://www.emailarchitect.net/webapp/download/easendmail.exe) executable from the provided link.
2. Install EA Sendmail Component Manager.

### *3.3 Updation of chromedriver*
1. Launch Chrome browser or install it if not already installed.
2. Navigate to Settings > About Chrome and note the version number.
3. Download the compatible version of Chromedriver [Download](https://googlechromelabs.github.io/chrome-for-testing/) using the provided link based on your Chrome version.
4. Replace the existing Chromedriver in the installation path of Selenium Basic with the downloaded one.

   ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/bd2d4878-bfd6-4da6-a3eb-e2df1992f48b)
</br>
</br>
</br>

# 4 Download Excel Sheet

### *4.1 Clone this repository [Or] Click Here to download* ⇊ [Download](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/releases/download/Agiletime2/AgileTime_v1.xlsm)

Clone this repository to your local machine to access the macro-enabled Excel file or download it from releases

1. Make a new folder and open the command prompt.
2. Input the below command.

```bash
git clone https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro.git
```

3. After completion, you will be able to access the macro-enabled Excel file.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/2ab0b2ed-b133-4609-b52f-0c219861c77b)


### *4.2 To grant security permission to the downloaded Excel sheet*

1. Right-click on the downloaded Excel file and select 'Properties.'
2. Ensure that the 'Unblock' option is checked under the general tab. Click 'Apply' and then 'OK' to enable security permissions for the Excel sheet, as shown in the screenshot below.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/dbe2a259-352a-44b3-b92c-a4bab26dbcbd)

</br>
</br>
</br>

# 5 Enabling custom macro-enabled buttons in an Excel sheet:

1. Open the downloaded AgileTime Excel file.
2. Navigate to File > Excel Options > Quick Access Toolbar. Check that the option 'Show Quick Access Toolbar below the Ribbon' is selected, then click 'OK' to apply the changes. as shown in the screenshot below.

![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/1b44c64a-a656-41f9-8351-e27af59d5b17)


4. Now, you will see the buttons ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/d9a1c4ef-7d5a-4f90-8700-f738127dacbd) enabled in the Excel sheet.

</br>
</br>
</br>

# 6 Features Overview

1. **User Details Sheet** - Input your Agile Time username, password, email ID, and manager ID in the corresponding cells.
2. **Master Details Sheet** and **'Update Master Data'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/14bb8ace-4570-4c32-8ad0-f7e5ad542efc)
 Button - Automatically updates eligible task details from your Agile Time records when you click 'Update master data.'
3. **Timesheet Details Sheet** and **'Upload Timesheet'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/7d4708ad-23b3-4cc1-b2a2-928084c2fd9b)
 Button - Fill out task details, clicking 'Upload timesheet' updates your timesheet in Agile Time.
4. **Configuration Sheet** and **'Mail My Timesheet'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/ac17dcb8-61b3-4206-be30-bbb287af1cd7)
 Button - Configure email settings like SMTP address, host, sender email ID, subject, and body. 'Mail my Timesheet' sends an email to your specified address.
5. **'User Details'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/bd02e629-c637-4a98-a04d-9ae038f5daed)
 Button - Navigates to the 'User Details' Sheet.
6. **'Master Details'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/89aa7427-2075-47d3-aa5f-e8d167a73b4b)
 Button - Navigates to the 'Master Details' Sheet.
7. **'Mail Configuration'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/d97b6c7f-b9d8-4be6-ba09-f3f7884d7664)
  Button - Navigates to the 'Configuration' Sheet.
8. **'Fill Timesheet Details'** ![image](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/assets/142989519/cc380356-1ef6-42f5-b873-1f35dbcb48af)
 Button - Navigates to the 'Timesheet Details' Sheet upon click.
</br>
</br>

### *Thank you for using Agile Timesheet Automation!*
*Report a problem(Issue) [Click Here](https://github.com/Jaisurya-Ravi/Agile-Time-Using-Excel-Macro/issues/new) when encountering difficulties with configuration or when using the functionality.*
