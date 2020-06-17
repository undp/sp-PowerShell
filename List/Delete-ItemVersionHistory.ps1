<#
	.SYNOPSIS
		Delete version history of a list item.

	.DESCRIPTION
		The Delete-ItemVersionHistory script deletes version histories of a list item given an upper and lower bound of version numbers.

	.PARAMETER SiteUrl
		The URL of the site.

	.PARAMETER ListName
		The list name that contains the item.

	.PARAMETER ItemId
		The item id in whicvh version histories will be deleted.

	.PARAMETER Start
		The version number to start deleting from.

	.PARAMETER End
		The version number to stop deleting at (inclusively).

	.EXAMPLE
		Delete-ItemVersionHistory.ps1 -SiteUrl $SiteUrl -List $List -ItemId $ItemId -Start $Start -End $End

	.EXAMPLE
		Delete-ItemVersionHistory.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MySite" -List "My List" -ItemId 3 -Start 5 -End 10

	.EXAMPLE
		Delete-ItemVersionHistory.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MySite" -List "My List" -ItemId 3 -Start 5 -End 5

	.NOTES
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>

param(
    [Parameter(Mandatory=$true)][string]$SiteUrl = $(throw "Please specify the URL of the site"),
	[Parameter(Mandatory=$true)][string]$List = $(throw "Please specify the list name"),
	[Parameter(Mandatory=$true)][string]$ItemId = $(throw "Please specify the item id"),
    [Parameter(Mandatory=$true)][string]$Start = $(throw "Please specify the start version"),
	[Parameter(Mandatory=$true)][string]$End = $(throw "Please specify the end version (inclusive)")
)

#region main
if ($Start -le $End -and $Start -gt 0) {
	try {
		Connect-PnPOnline -Url $SiteUrl -UseWebLogin
		$context = Get-PnPContext

		$item = Get-PnPListItem -List $List -Id $ItemId

		$versions = $item.Versions
		$context.Load($versions)
		$context.ExecuteQuery()

		$ctr = $versions.Count - 1
		$currentVersion = $versions[$ctr].VersionLabel -as [int]
		$deleted = 0

		while ($currentVersion -le $End -and $ctr -ge 0) {
			if ($currentVersion -ge $Start -and -not $versions[$ctr].IsCurrentVersion) {
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
#endregion
