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
# # 11 May 2020 Add deletion functionality, script completion

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
########################## MAIN ###############################################
if ($start -le $end -and $start -gt 0) {
	try {
		Connect-PnPOnline -Url $siteUrl -UseWebLogin
		$context = Get-PnPContext

		$item = Get-PnPListItem -List $list -Id $itemId

		$versions = $item.Versions
		$context.Load($versions)
		$context.ExecuteQuery()

		$ctr = $versions.Count - 1
		$currentVersion = $versions[$ctr].VersionLabel -as [int]
		$deleted = 0

		while ($currentVersion -le $end -and $ctr -ge 0) {
			if ($currentVersion -ge $start -and -not $versions[$ctr].IsCurrentVersion) {
				Write-Host $versions[$ctr].VersionLabel
				$versions[$ctr].DeleteObject()
				$context.ExecuteQuery()

				$deleted++
			}
			
			$currentVersion = $versions[--$ctr].VersionLabel -as [int]
		}

		Write-Host "Total versions deleted: ${deleted}"
	} catch {
		$exceptionName = $_.Exception.GetType().Name
		Write-Error $_
		Write-Host -ForegroundColor "red" "[${exceptionName}]"
	} finally {
		Disconnect-PnPOnline
	}
} else {
		Write-Host -ForegroundColor "red" "Start cannot be below 1 or greater than end."
	}  
#endregion ####################################################################