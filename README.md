# Microsoft-Office-2021
Microsoft Office 2021

# Table of Contents
   * [Download and install Office 2021](#download-and-install-office-2021)
      * [Install Office 2021 using Office Deployment Tool](#install-office-2021-using-office-deployment-tool)
   * [Activate Microsoft Office 2021](#activate-microsoft-office-2021)
      * [Method 1: Activate Microsoft Office 2021 using Command Prompt](#method-1-activate-microsoft-office-2021-using-command-prompt)
     
# Download and install Office 2021
## Install Office 2021 using Office Deployment Tool

Before you start, ensure that the operating system version you’re running is Windows 10 or later. There is no way to install Office 2021 on Windows 8.1 or earlier.

**Step 1:** Download the exe file from: [office 2021](https://c2rsetup.officeapps.live.com/c2r/download.aspx?ProductreleaseID=ProPlus2021Retail&platform=x64&language=en-us&version=O16GA)

Installing **Microsoft Office 2021**, takes several minutes, depending on your internet speed.

![Installing Office](https://github.com/user-attachments/assets/12152d10-0437-444b-9da0-7951e3c18c15)

Once you complete the installation, all Office apps will run under a 7-day trial license.



![Activation Required](https://github.com/user-attachments/assets/d7145ab7-13b4-43bb-b70c-3b3f7c0cb24e0-7951e3c18c15)


# Activate Microsoft Office 2021
## Method 1: Activate Microsoft Office 2021 using Command Prompt
**Step 1:** Type **cmd** in search box, right click on **Command Prompt** then select **Run as administrator**.

![Command-Prompt](https://github.com/user-attachments/assets/1cd4808a-ed18-46ef-8b92-7a1ae33cfa1e)

**Step 2:** Copy all below commands, **right click to paste into cmd window** at once then hit Enter.

```
if exist "C:\Program Files\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files\Microsoft Office\Office16"
if exist "C:\Program Files (x86)\Microsoft Office\Office16\ospp.vbs" cd /d "C:\Program Files (x86)\Microsoft Office\Office16"
for /f %x in ('dir /b ..\root\Licenses16\ProPlus2021VL_KMS*.xrm-ms') do cscript ospp.vbs /inslic:"..\root\Licenses16\%x"
cscript ospp.vbs /inpkey:FXYTK-NJJ8C-GB6DW-3DYQT-6F7TH
cscript ospp.vbs /sethst:kms.msgang.com
cscript ospp.vbs /act
pause
```
**Microsoft Office 2021** was activated, and you can start using the applications without any restrictions or limitations.

![Command-Prompt-Window](https://github.com/user-attachments/assets/ffde5317-3904-4a6b-9665-0aa3b7cc70b8)

**Product Activated**

![Image](https://github.com/user-attachments/assets/e9c9e21c-e967-482e-90ff-6a80d4ec6380)
