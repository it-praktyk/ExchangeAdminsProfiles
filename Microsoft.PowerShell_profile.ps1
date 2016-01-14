<#
    .SYNOPSIS
    PowerShell profile used by Exchange Server administrator
   
    More about PowerShell profiles 
    get-help about_Profiles
    https://technet.microsoft.com/en-us/library/hh847857.aspx
      
    .NOTES
    AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
    KEYWORDS: PowerShell, Exchange, Profiles
   
    VERSIONS HISTORY
    0.1.0 - 2015-03-08 - The first version publishid on GitHub
    0.2.0 - 2015-03-18 - console and buffer resizing corrected, verifying if known file extensions are displayed
						new PSDrive Scripts added
    0.3.0 - 2015-12-10 - Clear variables set temporary in profile, Set default parameters for Export-CSV cmdlet
    0.3.1 - 2015-12-22 - Assigning default parameters corrected, background colors for console corrected
    0.3.2 - 2016-01-14 - Corrected mistake in the variable usage, added removing existing PsDrives, rewrote set PSDefaultParameterValues

   DISCLAIMER
   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
   any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
   performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
   (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
   or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
   
#>


$pshost = get-host
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

#Create new PSDrives - Desktop and Scripts
$ProfileDrive = ((Get-Item env:APPDATA).Value).substring(0, 2)

$DesktopPath = $ProfileDrive + (Get-Item env:HOMEPATH).Value + '\Desktop'

If ( test-path -Path "Desktop:" ) {

	Set-Location -Path $DesktopPath

	Get-PSDrive "Desktop" | Remove-PSDrive

}

New-PSDrive -Name Desktop -PSProvider FileSystem -Root $DesktopPath | Out-Null

$ScriptsPath = $ProfileDrive + (Get-Item env:HOMEPATH).Value + '\Documents\Scripts'

If (Test-Path -Path $ScriptsPath -PathType Container) {

	If ( test-path -Path "Scripts:" ) {

		Set-Location -Path $ProfileDrive

		Get-PSDrive "Scripts" | Remove-PSDrive

	}
	    
    New-PSDrive -Name Scripts -PSProvider FileSystem -Root $ScriptsPath | Out-Null
    
}

Set-Location -Path  Desktop: -ErrorAction SilentlyContinue

#Disable certificate revocation checking - specially needed for Exchange Servers without internet access
$key1Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing"
$key1 = Get-ItemProperty -Path $key1Path -Name State


$key2Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
$key2 = Get-ItemProperty -Path $key2Path -Name CertificateRevocation

If ($key1.state -ne 146944) {
    
    Set-ItemProperty -Path $key1Path -Name State -Value 146944 -Force
    
}


If ($key2.CertificateRevocation -ne 0) {
    
    Set-ItemProperty -Path $key2Path -Name CertificateRevocation -Value 0 -Force
    
}


#Display extensions for known files
$key3Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Explorer\Advanced"
$key3 = Get-ItemProperty -Path $key3Path -Name HideFileExt

If ($key3.HideFileExt -ne 0) {
    
    Set-ItemProperty -Path $key3Path -Name HideFileExt -Value 0 -Force
    
}

#Remove previously set variables
$VariablesToRemove = "key1Path", "key2Path", "key3Path", "key1", "key2", "key3", "$VariablesToRemove"

$VariablesToRemove | ForEach-Object -Process {
    
    Remove-Variable -Name $_ -ErrorAction SilentlyContinue
    
}

#Assign default parameters values to some cmdlets - only works with Powershell 3.0 and newer :-/
if (([Version]$psversiontable.psversion).major -ge 3) {
    
	$DefaultParameterVaulesToAdd = @(@('Export-CSV:Delimiter';';'),@('Export-CSV:Encoding','UTF8'),@('Export-CSV:NoTypeInformation','$true'))
	
	$DefaultParameterVaulesToAddCount = $DefaultParameterVaulesToAdd.Lenght
	
	for($i=0; $i -lt $DefaultParameterVaulesToAddCount; $i++){
       
	   $CurrentParameterValue = $DefaultParameterVaulesToAdd[$i][0]
	   
	   If ( $PSDefaultParameterValues.Contains($CurrentParameterValue) ) {
	   
			$PSDefaultParameterValues.Remove($CurrentParameterValue)
			
			$PSDefaultParameterValues.Add($DefaultParameterVaulesToAdd[$i][0],$DefaultParameterVaulesToAdd[$i][1])
	   
	   }
	   
     }
    
}