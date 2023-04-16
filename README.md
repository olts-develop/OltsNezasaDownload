# OltsNezasaDownload
Powershell script to download Nezasa booking files

## Root folder
Create a folder, and place all files into this directory: PS1 and BAT files.

The BAT files are for manual testing. The PS1 files are called from the BAT files, and can be called from the Windows Task Scheduler.

Edit the batch files **nezasa_download_files_STG.bat** and **nezasa_download_files_PROD.bat**, and assign the appropriate AGENCY code supplied by Nezasa.

## Credentials-Files

Run once to create secure credential file, depending on the environment:

```Get-Credential | EXPORT-CLIXML "SecureCredentialsSTG.xml"```
```Get-Credential | EXPORT-CLIXML "SecureCredentialsPROD.xml"```

Enter the password supplied by Nezasa for the AGENCY.

Two files will be created:

SecureCredentialsPROD.xml
SecureCredentialsSTG.xml

The files will contain an encoded password for the PROD and STG environments that can only be used on this machine, and are referenced in the PS1 files.

## Prerequisites

Download and install Windows Powershell 7.x.x (pwsh.exe) and xmlstarlet:

```choco install -y powershell-core xmlstarlet```

## Directory structure

The following directory structure will be created when the Powershell fiels are run:

**prod**   <-  Directory for the production import
|_**files**   <-  the script will write files into this directory 
  |_**todo**   <-  Agent Online will move files from the **files** directory into this directory before starting the processing, so that new downloads from the scrip will not conflict with the import of the files
    |_**error**   <-  If an error occurs during the import, the files will be moved to this **directory**
	|_ **save**   <-  After import the files will be moved to dynamically created directory per **month**
	  |_**2023-04**
	  |_**2023-05**
	  |_ ...
|_**log**   <-   the log files for the past few days


**stg**   <-  Directory for the staging import
|_**files**   <-  the script will write files into this directory 
  |_**todo**   <-  Agent Online will move files from the **files** directory into this directory before starting the processing, so that new downloads from the scrip will not conflict with the import of the files
    |_**error**   <-  If an error occurs during the import, the files will be moved to this directory
	|_ **save**   <-  After import the files will be moved to dynamically created directory per month
	  |_**2023-04**
	  |_**2023-05**
	  |_ ...
|_**log**   <-   the log files for the past few days
 
## Task Scheduler
 
To run the scripts every minute in the windows task scheduler:
 
1) Download and install Windows Powershell 7.x.x: pwsh.exe
	```choco install -y powershell-core xmlstarlet```
	
	Installing xmlstarlet will allow converting JSON to XML.
	
2) Create a new task in the windows task scheduler:
	Command: pwsh.exe 
	Parameters: -File nezasa_download_files_STG.ps1 -AGENCY xxxxxxx
	Working directory: the Nezasa import root directory in which the PS1 files are.
	
3) The trigger can run every minute, but the trigger should be set up, to stop the job after 5 minutes, to prevent any endless loops in the PS1 code from eating up all the memory on the machine.