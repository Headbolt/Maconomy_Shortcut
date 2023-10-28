###############################################################################################################
#
# ABOUT THIS PROGRAM
#
#   Maconomy_Shortcut.ps1
#   https://github.com/Headbolt/Maconomy_Shortcut
#
#   This script was designed to create or delete specific Maconomy shortcuts
#
#	Intended use is in Microsoft Endpoint Manager, as an intunewin app, although it can be used manually
#
###############################################################################################################################################
#
#	Usage
#		Maconomy_Shortcut.ps1 [-Create | -Delete] -Name <shortcut name> -Address <server address> -Port <server port>
#								-DB <database name> -NewInst -Protocol <https protocol> -AppPath <Path to Maconomy App>
#
#			Note -NewInst is optional
#		eg. Maconomy_Shortcut.ps1 -Create -Name short.lnk -Address 'https://c.deltekenterprise.com' -Port 443 -DB live -NewInst -Protocol TLSv1.2 -AppPath 'C:\Program Files (x86)\Deltek Maconomy 22.101.6\Maconomy.exe'
#		eg. Maconomy_Shortcut.ps1 -Delete -Name shortcut.lnk
#
###############################################################################################################################################
#
# HISTORY
#
#   Version: 1.0 - 28/10/2023
#
#	28/10/2023 - V1.0 - Created by Headbolt
#
###############################################################################################################################################
#
#   DEFINE VARIABLES & READ IN PARAMETERS
#
###############################################################################################################################################
#
param (
	[switch]$Create,
	[switch]$Delete,
	[string]$Name,
 	[string]$Address,
 	[string]$Port,
 	[string]$DB,
	[switch]$NewInst,
	[string]$Protocols,
	[string]$AppPath
)
#
$global:ScriptVer="1.0" # Set ScriptVersion for logging
#
$global:LocalLogFilePath="$Env:WinDir\temp\" # Set LogFile Path
$global:ScriptName="Application | Download and Install" # Set ScriptName for logging
#
$global:Name=$Name # Pull Name into a Global Variable
$global:Address=$Address # Pull Name into a Global Variable
$global:Port=$Port # Pull Name into a Global Variable
$global:DB=$DB # Pull Name into a Global Variable
$global:Protocols=$Protocols # Pull Name into a Global Variable
$global:AppPath=$AppPath # Pull Name into a Global Variable
#
###############################################################################################################################################
#
#   Functions Definition
#
###############################################################################################################################################
#
#   Logging Function
#
function Logging
{
#	
If ( $Create )
{
	$LocalLogFileType="_Create.log" # Set ActionType for Log File Path
	$global:LocalLogFilePath=$global:LocalLogFilePath+$global:Name+$LocalLogFileType # Construct Log File Path
}
#
If ( $Delete )
{
	$LocalLogFileType="_Delete.log" # Set ActionType for Log File Path
	$global:LocalLogFilePath=$global:LocalLogFilePath+$global:Name+$LocalLogFileType # Construct Log File Path
}
#
Start-Transcript $global:LocalLogFilePath # Start the logging
Clear-Host # Clear Screen
SectionEnd
Write-Host "Logging to $global:LocalLogFilePath"
Write-Host ''# Outputting a Blank Line for Reporting Purposes
Write-Host "Script Version $global:ScriptVer"
}     
#
###############################################################################################################################################
#
# Creation Function
#
Function Creation {
	$global:NewShortcut="c:\Users\Public\Desktop\$global:Name" # Setting the new Shortcut Path
	$global:TargetFile = "C:\Program Files (x86)\Deltek Maconomy 22.101.6\Maconomy.exe" # Set Path to the Application
	Write-Host 'Creating Shortcut "'$global:NewShortcut' "'
	$WScriptShell = New-Object -ComObject WScript.Shell
	$Shortcut = $WScriptShell.CreateShortcut($global:NewShortcut)
	Write-Host 'Targetting at EXE "'$global:AppPath' "'
	$Shortcut.TargetPath = $global:AppPath
	If ( $NewInst )
	{
		$global:Arguments = "-a $Address -p $Port -s $DB --new-instance -vmargs -Duser.home=%userprofile% -Dhttps.protocols=$global:Protocols" # Set Arguments required
	}
	else
	{
		$global:Arguments = "-a $Address -p $Port -s $DB -vmargs -Duser.home=%userprofile% -Dhttps.protocols=$global:Protocols" # Set Arguments required
	}
	Write-Host 'Using Arguments "'$global:Arguments' "'
	$Shortcut.Arguments = $global:Arguments 
	$Shortcut.Save()
}
#
###############################################################################################################
#
# Deletion Function
#
Function Deletion {
	Write-Host 'Attempting to Delete file "'c:\Users\Public\Desktop\$global:Name'"'
	Remove-Item -Path c:\Users\Public\Desktop\$global:Name -Force -erroraction 'silentlycontinue'
}
#
#
###############################################################################################################################################
#
# Section End Function
#
function SectionEnd
{
#
Write-Host '' # Outputting a Blank Line for Reporting Purposes
Write-Host  '-----------------------------------------------' # Outputting a Dotted Line for Reporting Purposes
Write-Host '' # Outputting a Blank Line for Reporting Purposes
#
}
#
###############################################################################################################################################
#
# Script End Function
#
Function ScriptEnd
{
#
Write-Host Ending Script $global:ScriptName
Write-Host '' # Outputting a Blank Line for Reporting Purposes
Write-Host  '-----------------------------------------------' # Outputting a Dotted Line for Reporting Purposes
Write-Host ''# Outputting a Blank Line for Reporting Purposes
#
Stop-Transcript # Stop Logging
#
}
#
###############################################################################################################################################
#
# End Of Function Definition
#
###############################################################################################################################################
# 
# Begin Processing
#
###############################################################################################################################################
#
Logging
#
SectionEnd
#
If ( $Create )
{
	Write-Host '"Create" action requested'
	SectionEnd
	If ( $Global:Name ) # Check Name is set
	{
#		Write-Host 'Name Variable set to "'$global:Name' "'
		If ( $Global:Address ) # Check Address is set
		{
#			Write-Host 'Address Variable set to "'$Global:Address' "'
			If ( $Global:Port ) # Check Port is set
			{
#				Write-Host 'Port Variable set to "'$Global:Port' "'
				If ( $Global:DB ) # Check DB is set
				{	
#					Write-Host 'DB Variable set to "'$Global:DB' "'
					If ( $Global:Protocols ) # Check Protocols is set
					{		
#						Write-Host 'Protocols Variable set to "'$Global:Protocols' "'
						If ( $Global:AppPath ) # Check AppPath is set
						{		
#							Write-Host 'AppPath Variable set to "'$Global:AppPath' "'
							Creation
						}
						Else
						{
							Write-Host 'AppPath Variable not set, Creation cannot continue'
						}
					}
					Else
					{
						Write-Host 'Protocols Variable not set, Creation cannot continue'
					}
				}
				Else
				{
					Write-Host 'DB Variable not set, Creation cannot continue'
				}
			}
			Else
			{
				Write-Host 'Port Variable not set, Creation cannot continue'
			}
		}
		Else
		{
			Write-Host 'Address Variable not set, Creation cannot continue'
		}
	}
	Else
	{
		Write-Host 'Name Variable not set, Creation cannot continue'
	}
}
#
If ( $Delete )
{
	Write-Host '"Delete" action requested'
	SectionEnd
	If ($Global:Name ) # Check Name  is set
	{
		Deletion
	}
}
#
SectionEnd
ScriptEnd
