<#
	.SYNOPSIS
		Deletes version histories of an item.

	.DESCRIPTION
		Deletes version histories of an item given an upper and lower bound of version numbers.

	.PARAMETER siteUrl
		The url of the site.

	.PARAMETER itemUrl
		The item url in which version histories will be deleted.

	.PARAMETER start
		The version number to start deleting from.

	.PARAMETER end
		The version number to stop deleting at (inclusively).

	.INPUTS
		System.String,System.String,System.Int32,System.Int32

	.EXAMPLE
		Delete-VersionHistories -siteUrl $siteUrl -itemUrl $itemUrl -start $start -end $end

	.EXAMPLE
		Delete-VersionHistories -siteUrl "https://tenant.sharepoint.com/Sites/mysite" -itemUrl "/Shared Documents/mydoc.txt" -start 5 -end 10

	.EXAMPLE
		Delete-VersionHistories -siteUrl "https://tenant.sharepoint.com/Sites/mysite" -itemUrl "/Shared Documents/mydoc.txt" -start 5 -end 5

	.NOTES
AUTHOR: Giuseppe Campanelli
LASTEDIT: $(Get-Date)
# Written by: Giuseppe Campanelli (giuseppe.campanelli@undp.org) on 8 April 2020
#
# version: 1.0
#
# Revision history
# # 08 Apr 2020 Creation

#>


#region command line parameters ###############################################
# Input Parameters
# $args: url of the site, url of the item, start version, end version
param(
    [Parameter(Mandatory=$true)][string]$siteUrl = $(throw "Please specify the URL for site for the item"),
	[Parameter(Mandatory=$true)][string]$itemUrl = $(throw "Please specify the URL for item"),
    [Parameter(Mandatory=$true)][string]$start = $(throw "Please specify the start version"),
	[Parameter(Mandatory=$true)][string]$end = $(throw "Please specify the end version (inclusive)")
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