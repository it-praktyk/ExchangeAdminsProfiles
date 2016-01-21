# ExchangeAdminsProfiles
PowerShell profiles with settings helpfull for Exchange Server adminstrators

## Included settings
- additional PSDrive 'Desktop:' which pointing to Desktop
- additional PSDrive 'Scripts:' whic pointing to MyDocuments\Scripts - if the folder exist
- switching off the certificates revocation list (the same as Advanced Options ) - it is specially helpfull for Exchange Servers which don't have access to the internet - Exchange Management Shell is starting faster
- console resizing and color standarization
- adding default parameters settings to the variable PSDefaultParameterValues - for Export-CSV e.g.
- display file extensions for known file format

	
 
## More about PowerShell profiles 
- get-help about_Profiles
- https://technet.microsoft.com/en-us/library/hh847857.aspx
      
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

##DISCLAIMER
   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
   any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
   performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
   (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
   or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 