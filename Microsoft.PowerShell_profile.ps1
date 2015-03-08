<#
  .SYNOPSIS
   PowerShell profile used by Exchange Server administrator
   
   More about PowerShell profiles 
   get-help about_Profiles
   https://technet.microsoft.com/en-us/library/hh847857.aspx
      
  .NOTES
   AUTHOR: Wojciech Sciesinski, wojciech[at]sciesinski[dot]net
   KEYWORDS: PowerShell, Profiles
   
   VERSIONS HISTORY
   0.1.0 -  2015-03-08 - The first version publishid on GitHub

   TODO
   Verify/change settings windows/buffer size 
   
   DISCLAIMER
   This script is provided AS IS without warranty of any kind. I disclaim all implied warranties including, without limitation,
   any implied warranties of merchantability or of fitness for a particular purpose. The entire risk arising out of the use or
   performance of the sample scripts and documentation remains with you. In no event shall I be liable for any damages whatsoever
   (including, without limitation, damages for loss of business profits, business interruption, loss of business information,
   or other pecuniary loss) arising out of the use of or inability to use the script or documentation. 
   
#>


	$pshost = get-host
	$console = $pshost.UI.RawUI

	$foreground = "black"

	$console.ForegroundColor = "white"
	$console.BackgroundColor = $foreground

	$console2 = (Get-Host).PrivateData

	$console2.ErrorForegroundColor = "red"
	$console2.ErrorBackgroundColor = $foreground
	$console2.WarningForegroundColor = "yellow"
	$console2.WarningBackgroundColor = $foreground
	$console2.DebugForegroundColor = "yellow"
	$console2.DebugBackgroundColor  = $foreground
	$console2.VerboseForegroundColor = "yellow"
	$console2.VerboseBackgroundColor  = $foreground
	$console2.ProgressForegroundColor  = "yellow"
	$console2.ProgressBackgroundColor = "darkcyan"

	$buffer = $console.BufferSize
	$buffer.Width = 120
	$buffer.Height = 1000
	$console.BufferSize = $buffer


	$size = $console.WindowSize
	$size.Width = 120
	$size.Height = 50
	$console.WindowSize = $size

	$DesktopPath= 'c:\' + (Get-Item env:HOMEPATH).Value +'\Desktop'
	New-PSDrive -Name Desktop -PSProvider FileSystem -Root $DesktopPath | Out-Null

	cd Desktop:

	#Disable certificate revocation checking - specially needed for Exchange Servers without internet access
	$key1Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\WinTrust\Trust Providers\Software Publishing"
	$key1 = Get-ItemProperty -Path $key1Path -Name State 


	$key2Path = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Internet Settings"
	$key2 = Get-ItemProperty -Path $key2Path  -Name CertificateRevocation

	If ($key1.state -ne 146944) {

		Set-ItemProperty -Path $key1Path -Name State -Value 146944 -Force

	}


	If ($key2.CertificateRevocation -ne 0) {

		Set-ItemProperty -Path $key2Path -Name CertificateRevocation -Value 0 -Force

	}

	Clear-Host