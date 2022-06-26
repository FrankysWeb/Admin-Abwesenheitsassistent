#--------------------------------------------
# Declare Global Variables and Functions here
#--------------------------------------------


#Sample function that provides the location of the script
function Get-ScriptDirectory
{
<#
	.SYNOPSIS
		Get-ScriptDirectory returns the proper location of the script.

	.OUTPUTS
		System.String
	
	.NOTES
		Returns the correct path within a packaged executable.
#>
	[OutputType([string])]
	param ()
	if ($null -ne $hostinvocation)
	{
		Split-Path $hostinvocation.MyCommand.path
	}
	else
	{
		Split-Path $script:MyInvocation.MyCommand.Path
	}
}

#Sample variable that provides the location of the script
[string]$ScriptDirectory = Get-ScriptDirectory

#Load Exchange Shell
try
{
	Add-PSSnapin Microsoft.Exchange.Management.PowerShell.SnapIn
}
catch
{
	Write-Host "Failed to Load Exchange Shell"
	break
}

#Internal OOF Message
$InternalMessage = "Hallo Kollegen,

ich bin bis zum XX.XX.XXXX nicht im Haus und bearbeite Mails daher nicht.
Ihre Mail wird nicht automatisch weitergeleitet. Bitte wenden Sie sich in dringenden Fällen an:

Vertretung XX

Mit frerundlichen Grüßen,

Vorname Nachname"

#External OOF Message
$ExternalMessage = "Sehr geehrte Damen und Herren,

ich bin bis zum XX.XX.XXXX nicht im Haus und bearbeite Mails daher nicht.
Ihre Mail wird nicht automatisch weitergeleitet. Bitte wenden Sie sich in dringenden Fällen an:

Vertretung XX

Mit frerundlichen Grüßen,

Vorname Nachname"