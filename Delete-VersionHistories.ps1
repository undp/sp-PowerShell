<#
	.SYNOPSIS
		Deletes version histories of a list item.

	.DESCRIPTION
		Deletes version histories of a list item given an upper and lower bound of version numbers.

	.PARAMETER siteUrl
		The url of the site.

	.PARAMETER list
		The list name that contains the item.

	.PARAMETER itemId
		The item id in whicvh version histories will be deleted.

	.PARAMETER start
		The version number to start deleting from.

	.PARAMETER end
		The version number to stop deleting at (inclusively).

	.INPUTS
		System.String,System.String,System.Int32,System.Int32

	.EXAMPLE
		Delete-VersionHistories -siteUrl $siteUrl -list $list -itemId $itemId -start $start -end $end

	.EXAMPLE
		Delete-VersionHistories -siteUrl "https://tenant.sharepoint.com/Sites/mysite" -list "My List" -itemId 3 -start 5 -end 10

	.EXAMPLE
		Delete-VersionHistories -siteUrl "https://tenant.sharepoint.com/Sites/mysite" -list "My List" -itemId 3 -start 5 -end 5

	.NOTES
AUTHOR: Giuseppe Campanelli
LASTEDIT: $(Get-Date)
# Written by: Giuseppe Campanelli (giuseppe.campanelli@undp.org) on 8 April 2020
#
# version: 1.0
#
# Revision history
# # 08 Apr 2020 Creation
# # 06 May 2020 Update input parameters, retrieve versions of list item

#>


#region command line parameters ###############################################
# Input Parameters
# $args: url of the site, url of the item, start version, end version
param(
    [Parameter(Mandatory=$true)][string]$siteUrl = $(throw "Please specify the URL of the site"),
	[Parameter(Mandatory=$true)][string]$list = $(throw "Please specify the list name"),
	[Parameter(Mandatory=$true)][string]$itemId = $(throw "Please specify the item id"),
    [Parameter(Mandatory=$true)][string]$start = $(throw "Please specify the start version"),
	[Parameter(Mandatory=$true)][string]$end = $(throw "Please specify the end version (inclusive)")
	);
#endregion ####################################################################

#region MAIN ##################################################################
########################## MAIN #####################################
if ($start -le $end -and $start -gt 0) {
	try {
		Connect-PnPOnline -Url $siteUrl -UseWebLogin
		$context = Get-PnPContext

		$item = Get-PnPListItem -List $list -Id $itemId

		$versions = $item.Versions
		$context.Load($versions)
		$context.ExecuteQuery()

		$amountToDelete = 1 + $end - $start

		if ($end -le $versions.Count) {
			while ($amountToDelete -gt 0 -and $versions.Count -gt 1) {
				Write-Host $versions[$versions.Count - $start].VersionLabel
				$amountToDelete--
			}
		} else {
			Write-Host -ForegroundColor "red" "Please ensure the start and end is a valid range."
		}
	} catch {
		$exceptionName = $_.Exception.GetType().Name
		Write-Host -ForegroundColor "red" "[${exceptionName}]"
	} finally {
		Disconnect-PnPOnline
	}
} else {
		Write-Host -ForegroundColor "red" "Start cannot be below 1 or greater than end."
	}
#endregion ####################################################################