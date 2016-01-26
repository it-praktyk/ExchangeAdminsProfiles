# ExchangeAdminsProfiles
PowerShell profiles with settings helpfull for Exchange servers adminstrators

## Included settings
- console resizing and color standarization
- additional PSDrive 'Desktop:' which pointing to Desktop
- additional PSDrive 'Scripts:' which pointing to MyDocuments\Scripts - if the folder exist
- switching off the certificates revocation list (the same as in Internet Explorer -> Options -> Advanced ) - it is specially helpfully for Exchange Servers which don't have access to the internet - Exchange Management Shell is starting faster
- display file extensions for known file format
- adding default parameters settings to the variable PSDefaultParameterValues - for Export-CSV e.g.


## More about PowerShell profiles
- get-help about_Profiles
- https://technet.microsoft.com/en-us/library/hh847857.aspx

	Copy text from this file to the PowerShell profile file (create if doesn't exist)
    - for you only: C:\&lt;YOUR_PROFILE_DIRECTORY&gt;\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1 - value for the variable $Profile.CurrentUserCurrentHost
    - for all users on computer: C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1 - value for the variable $Profile.AllUsersCurrentHost

## EXAMPLES

### EXAMPLE 1

```powershell

    You can check path for the profiles examining the $profile variable

    [PS] > $Profile | Format-List -Force

    AllUsersAllHosts       : C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
    AllUsersCurrentHost    : C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1
    CurrentUserAllHosts    : C:\Users\Username\Documents\WindowsPowerShell\profile.ps1
    CurrentUserCurrentHost : C:\Users\UserName\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
    Length                 : 76
```

### EXAMPLE 2

```powershell
You can examine if the profile file exist if and create empty one a a profile file if not exist

[PS] > If (!(Test-Path -Path $Profile)) { New-Item -Path $Profile -ItemType File -Force }

        Directory: C:\Users\Wojtek\Documents\WindowsPowerShell


        Mode                LastWriteTime         Length Name
        ----                -------------         ------ ----
        -a----        1/22/2016  12:01 AM              0 Microsoft.PowerShell_profile.ps1

The empty file will be created which will be used as a profile for the current user only.

For all users use the variable $Profile.AllUsersAllHosts instead $Profile .
```


## AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net  

## KEYWORDS: PowerShell, Exchange, Profiles

## VERSIONS HISTORY
    - 0.1.0 - 2015-03-08 - The first version publishid on GitHub
    - 0.2.0 - 2015-03-18 - console and buffer resizing corrected, verifying if known file extensions are displayed
						new PSDrive Scripts added
    - 0.3.0 - 2015-12-10 - Clear variables set temporary in profile, Set default parameters for Export-CSV cmdlet
    - 0.3.1 - 2015-12-22 - Assigning default parameters corrected, background colors for console corrected
    - 0.3.2 - 2016-01-14 - Corrected mistake in the variable usage, added removing existing PsDrives, rewrote set PSDefaultParameterValues
	- 0.3.3 - 2016-01-21 - Removing temporary variables corrected
	- 0.4.0 - 2016-01-21 - The file reformated, help updated
	- 0.5.0 - 2016-01-25 - Set registry rewrote, set PSDefaultParameterValues corrected
	- 0.6.0 - 2016-01-26 - Set PSDrive rewrote

## DISCLAIMER
   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
   any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
   performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
   (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
   or other pecuniary loss) arising out of the use of or inability to use the script or documentation.
