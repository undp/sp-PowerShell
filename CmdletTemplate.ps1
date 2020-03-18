<#
	.SYNOPSIS
	A brief description of the function or script. This keyword can be used only once in each topic.
	Adds a file name extension to a supplied name.

	.DESCRIPTION
	A detailed description of the function or script. This keyword can be used only once in each topic.
	Adds a file name extension to a supplied name. Takes any strings for the file name or extension.

	.PARAMETER sourcewebURL
	The description of a parameter. Add a ".PARAMETER" keyword for each parameter in the function or script syntax.
	Specifies the Source SPWeb URL (note the parameter named above should match the parameter name in the parameter block if you want this help message to be displayed in Get-Help with -Full.

	.PARAMETER destwebURL
	Specifies the destination SPWebl URL. "Txt" is the default.

	.PARAMETER listTitle
	Specifies the SPList Title. Use the DisplayName, not the internal listname

	.PARAMETER backupPath
	Specifies the path for the backup file

	.INPUTS
	The Microsoft .NET Framework types of objects that can be piped to the function or script. You can also include a description of the input objects.
	None. You cannot pipe objects to Add-Extension.

	.OUTPUTS
	The .NET Framework type of the objects that the cmdlet returns. You can also include a description of the returned objects.
	System.String. Add-Extension returns a string with the extension
	or file name.

	.EXAMPLE
	A sample command that uses the function or script, optionally followed by sample output and a description. Repeat this keyword for each example.
	C:\PS> extension -name "File"
	File.txt

	.EXAMPLE
	C:\PS> extension -name "File" -extension "doc"
	File.doc

	.EXAMPLE
	C:\PS> extension "File" "doc"
	File.doc

	.LINK
	The name of a related topic. The value appears on the line below the ".LINK" keyword and must be preceded by a comment symbol # or included in the comment block.
	http://www.fabrikam.com/extension.html

	.LINK
	Set-Item

	.NOTES
AUTHOR: Daniel Tshin
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 11 June 2019
#
# version: 1.4
#
# Revision history
# # 11 Jun 2019 Creation

#>


#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space
param(
    [Parameter(Mandatory=$true)][string]$sourcewebURL = $(throw "Please specify the URL for the source site for the list(s)"),
	[Parameter(Mandatory=$true)][string]$destwebURL = $(throw "Please specify the URL for the destination site for the list(s)"),
    [Parameter(Mandatory=$true)][string]$listTitle = $(throw "Please specify the Title(s) of the list(s), separated by commas"),
	[Parameter(Mandatory=$false)][string]$backupPath
	);
#endregion ####################################################################

Process
{
	# For MS Teams
	Import-Module MicrosoftTeams
	Import-Module AzureAD
	#$cred = Get-Credential
	#$username = $cred.UserName
	Connect-MicrosoftTeams #-Credential $cred
	Connect-AzureAD
	# End MS Teams

	# For SPO
	Connect-PnPOnline -Url $sourcewebURL -UseWebLogin
	# End SPO

	Try {

	}
	Catch {
		write-host -f Red "Error in Script: " $_.Exception.Message
	}
}