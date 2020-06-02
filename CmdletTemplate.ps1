<#
	.SYNOPSIS
		A brief description of the function or script. This keyword can be used only once in each topic.
		Adds a file name extension to a supplied name.

	.DESCRIPTION
		A detailed description of the function or script. This keyword can be used only once in each topic.
		Adds a file name extension to a supplied name. Takes any strings for the file name or extension.

	.PARAMETER WebUrl
		The description of a parameter. Add a ".PARAMETER" keyword for each parameter in the function or script syntax.
		Specifies the Source SPWeb URL (note the parameter named above should match the parameter name in the parameter block if you want this help message to be displayed in Get-Help with -Full.

	.PARAMETER ListTitle
		Specifies the SPList Title. Use the DisplayName, not the internal listname

	.INPUTS
		The Microsoft .NET Framework types of objects that can be piped to the function or script. You can also include a description of the input objects.
		None. You cannot pipe objects to Add-Extension.

	.OUTPUTS
		The .NET Framework type of the objects that the cmdlet returns. You can also include a description of the returned objects.
		System.String. Add-Extension returns a string with the extension or file name.

	.EXAMPLE
		A sample command that uses the function or script, optionally followed by sample output and a description. Repeat this keyword for each example.
		CmdletTemplate.ps1 -WebUrl $WebUrl -ListTitle $ListTitle

	.EXAMPLE
		CmdletTemplate.ps1 -WebUrl "https://tenant.sharepoint.com/sites/MySite" -ListTitle "MyList"

	.LINK
		The name of a related topic. The value appears on the line below the ".LINK" keyword and must be preceded by a comment symbol # or included in the comment block.
		http://www.fabrikam.com/extension.html

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
#>

param(
    [Parameter(Mandatory=$true)][string]$WebUrl = $(throw "Please specify the URL of site for the list."),
    [Parameter(Mandatory=$true)][string]$ListTitle = $(throw "Please specify the name of the list.")
)

#region main
Begin {}
Process
{
	# For MS Teams
	#Import-Module MicrosoftTeams
	#Import-Module AzureAD
	#$cred = Get-Credential
	#$username = $cred.UserName
	#Connect-MicrosoftTeams #-Credential $cred
	#Connect-AzureAD
	# End MS Teams

	# For SPO
	Connect-PnPOnline -Url $WebUrl -UseWebLogin
	# End SPO

	try {
		#Get List by name
		#Return/output list count
	} catch {
		Write-Host -f Red "Error in Script: " $_.Exception.Message
	} finally {
		Disconnect-PnPOnline
	}
}
End {}
#endregion
