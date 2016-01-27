<#
    .SYNOPSIS
    PowerShell profiles with settings helpfull for Exchange Server adminstrators
   
    .DESCRIPTION
    PowerShell profiles with settings helpfull for Exchange Server adminstrators
	
	Included settings
	- additional PSDrive 'Desktop:' which pointing to Desktop
	- additional PSDrive 'Scripts:' whic pointing to MyDocuments\Scripts - if the folder exist
	- switching off the certificates revocation list (the same as Advanced Options ) - it is specially helpfull for Exchange Servers which don't have access to the internet - Exchange Management Shell is starting faster
	- console resizing and color standarization
	- adding default parameters settings to the variable PSDefaultParameterValues - for Export-CSV e.g.
	- display file extensions for known file format
    
    More about PowerShell profiles 
    get-help about_Profiles
    https://technet.microsoft.com/en-us/library/hh847857.aspx

    Copy text from this file to the file (create if doesn't exist)
    - for you only: C:\<YOUR_PROFILE_DIRECTORY>\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1 - value for the variable $Profile.CurrentUserCurrentHost
    - for all users on computer: C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1 - value for the variable $Profile.AllUsersCurrentHost

    .EXAMPLE
    You can check path for the profiles examining the $profile variable

    [PS] > $Profile | Format-List -Force

    AllUsersAllHosts       : C:\Windows\System32\WindowsPowerShell\v1.0\profile.ps1
    AllUsersCurrentHost    : C:\Windows\System32\WindowsPowerShell\v1.0\Microsoft.PowerShell_profile.ps1
    CurrentUserAllHosts    : C:\Users\Username\Documents\WindowsPowerShell\profile.ps1
    CurrentUserCurrentHost : C:\Users\UserName\Documents\WindowsPowerShell\Microsoft.PowerShell_profile.ps1
    Length                 : 76

    .EXAMPLE
    You can examine if the profile file exist if and create empty one a a profile file if not exist
    
    [PS] > If (!(Test-Path -Path $Profile)) { New-Item -Path $Profile -ItemType File -Force }

        Directory: C:\Users\Wojtek\Documents\WindowsPowerShell


        Mode                LastWriteTime         Length Name
        ----                -------------         ------ ----
        -a----        1/22/2016  12:01 AM              0 Microsoft.PowerShell_profile.ps1

    The empty file will be created which will be used as a profile for the current user only.

    For all users use the variable $Profile.AllUsersAllHosts instead $Profile .
      
    .LINK
    https://github.com/it-praktyk/ExchangeAdminsProfiles
    
    .LINK
    https://www.linkedin.com/in/sciesinskiwojciech
    
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, Profiles
       
    VERSIONS HISTORY
    - 0.1.0 - 2015-03-08 - The first version publishid on GitHub
    - 0.2.0 - 2015-03-18 - console and buffer resizing corrected, verifying if known file extensions are displayed new PSDrive Scripts added
    - 0.3.0 - 2015-12-10 - Clear variables set temporary in profile, Set default parameters for Export-CSV cmdlet
    - 0.3.1 - 2015-12-22 - Assigning default parameters corrected, background colors for console corrected
    - 0.3.2 - 2016-01-14 - Corrected mistake in the variable usage, added removing existing PsDrives, rewrote set PSDefaultParameterValues
    - 0.3.3 - 2016-01-21 - Removing temporary variables corrected
    - 0.4.0 - 2016-01-21 - The file reformated, help updated
    - 0.5.0 - 2016-01-25 - Set registry rewrote, set PSDefaultParameterValues corrected
	- 0.6.0 - 2016-01-26 - Set PSDrive rewrote
	- 0.6.1 - 2016-01-27 - Help updated

    TODO
    - create script for install profile for the local and remote computer
    - create script for (automatic) update (with merging ?) profile
    
    DISCLAIMER
    This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
    any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
    performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
    (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
    or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
   
#>

$pshost = get-host

#Resize and set other atributes only for console not for ISE environment
If ($pshost.Name -eq 'ConsoleHost') {
    
    $console = $pshost.UI.RawUI
    
    # All available colors you can chech using - source: http://blogs.technet.com/b/gary/archive/2013/11/21/sample-all-powershell-console-colors.aspx
    # [enum]::GetValues([System.ConsoleColor]) | Foreach-Object {Write-Host $_ -ForegroundColor $_}    
    
    $BackgroundColor = "black"
    
    $console.ForegroundColor = "white"
    $console.BackgroundColor = $BackgroundColor
    
    $console2 = (Get-Host).PrivateData
    
    $console2.ErrorForegroundColor = "red"
    $console2.ErrorBackgroundColor = $BackgroundColor
    $console2.WarningForegroundColor = "yellow"
    $console2.WarningBackgroundColor = $BackgroundColor
    $console2.DebugForegroundColor = "yellow"
    $console2.DebugBackgroundColor = $BackgroundColor
    $console2.VerboseForegroundColor = "yellow"
    $console2.VerboseBackgroundColor = $BackgroundColor
    $console2.ProgressForegroundColor = "yellow"
    $console2.ProgressBackgroundColor = $BackgroundColor
    
    $windowsSizeWidth = 120
    $windowsSizeHeight = 50
    $bufferSizeWidth = $WindowsSizeWidth
    $bufferSizeHeight = 1000
    
    Try {
        
        [system.console]::WindowHeight = $windowsSizeHeight
        [system.console]::WindowWidth = $windowsSizeWidth
        
        [system.console]::BufferHeight = $bufferSizeHeight
        [system.console]::BufferWidth = $bufferSizeWidth
        
    }
    Catch {
        
        [system.console]::WindowHeight = $windowsSizeHeight
        [system.console]::WindowWidth = 150
    }
    
}

#Create new PSDrives
$ProfileDrive = ((Get-Item env:APPDATA).Value).substring(0, 2)

$DesktopPath = $ProfileDrive + (Get-Item env:HOMEPATH).Value + '\Desktop'
$PSDrive1 = @("Desktop:","$DesktopPath")

$ScriptsPath = $ProfileDrive + (Get-Item env:HOMEPATH).Value + '\Documents\Scripts'
$PSDrive2 = @("Scripts:","$ScriptsPath")

$PSDrivesToCreate = @($PSDrive1,$PSDrive2)

for ($i = 0; $i -lt $PSDrivesToCreate.Length; $i++) {

	If (Test-Path -Path $PsDrivesToCreate[$i][1] -PathType Container) {

		If (test-path -Path ($PsDrivesToCreate[$i][0]))   {
        
			Set-Location -Path $ProfileDrive
        
			Get-PSDrive $($PsDrivesToCreate[$i][0]).Replace(':','')  | Remove-PSDrive -ErrorAction SilentlyContinue
        
		}	
	
		New-PSDrive -Name $($PsDrivesToCreate[$i][0]).Replace(':','') -PSProvider FileSystem -Root ($PsDrivesToCreate[$i][1]) | Out-Null
    
	}

}

Set-Location -Path Desktop: -ErrorAction SilentlyContinue

#Disable certificate revocation checking - specially needed for Exchange Servers without internet access
$RegistryKey1 = @("HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing", "State", "146944")
$RegistryKey2 = @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings", "CertificateRevocation", "0")

#Display extensions for known files
$RegistryKey3 = @("HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced", "HideFileExt", "0")

$RegistryKeys = @($RegistryKey1, $RegistryKey2, $RegistryKey3)

for ($i = 0; $i -lt ($RegistryKeys).Length; $i++) {
    
    if ((Get-ItemProperty -Path ($RegistryKeys[$i][0])).($RegistryKeys[$i][1]) -ne ($RegistryKeys[$i][2])) {
        
        Set-ItemProperty -Path $($RegistryKeys[$i][0]) -Name $($RegistryKeys[$i][1]) -Value $($RegistryKeys[$i][2]) -Force
        
    }
    
}


#Assign default parameters values to some cmdlets - only works with Powershell 3.0 and newer :-/
if (([Version]$psversiontable.psversion).major -ge 3) {
    
    $DefaultParameterVaulesToAdd = @(@('Export-CSV:Delimiter'; ';'), `
    @('Export-CSV:Encoding', 'UTF8'), `
    @('Export-CSV:NoTypeInformation', '$true'))
    
    for ($i = 0; $i -lt ($DefaultParameterVaulesToAdd).Lenght; $i++) {
        
        $CurrentParameterValue = $DefaultParameterVaulesToAdd[$i][0]
        
        If ($PSDefaultParameterValues.Contains($CurrentParameterValue)) {
            
            $PSDefaultParameterValues.Remove($CurrentParameterValue)
            
        }
        
        $PSDefaultParameterValues.Add($DefaultParameterVaulesToAdd[$i][0], $DefaultParameterVaulesToAdd[$i][1])
        
    }
    
}

#Remove previously set variables - please use the parameters names without "$" char
$VariablesToRemove = "PSDrive1", "PSDrive2", "PsDrivesToCreate", "RegistryKey1", "RegistryKey2", `
"RegistryKey3", "RegistryKeys", "DefaultParameterVaulesToAdd", "VariablesToRemove"

$VariablesToRemove | ForEach-Object -Process {
    
    Remove-Variable -Name $_ -ErrorAction SilentlyContinue
    
}