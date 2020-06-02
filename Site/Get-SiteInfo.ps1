<#
	.SYNOPSIS
		Get information about SharePoint Online site.

	.DESCRIPTION
		The Get-SiteInfo script gets information of SharePoint Online sites including Title, URL, WebTemplate, List Count, Subwebs Count, MasterPageURL, InheritsMasterURL, InheritsCustomMasterURL, Owner, Size, Language, Locale, ParentWebId, SiteLogoPath.
		Get-SiteInfo uses SharePoint-PnPPowerShell.

	.PARAMETER SiteURL
		Site url of the webs to inspect.
	
	.PARAMETER CSVFilePath
		CSV file path containing Site URLs of the webs to inspect.

	.EXAMPLE
		Get-SiteInfo.ps1 -SiteUrl "https://tenant.sharepoint.com/sites/MySite"

	.EXAMPLE
		Get-SiteInfo.ps1 -CSVFilePath "C:/CSVFolder/MyCSVFile.csv"

	.INPUTS
		System.String,System.String

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, suprememilanfan@gmail.com
#>

param(
	[Parameter(Mandatory=$false)][string]$SiteUrl,
	[Parameter(Mandatory=$false)][string]$CSVFilePath
)

if ($psboundparameters.Count -ne 1) {
	Write-Host "Pleasy specify only one of two parameters: -SiteURl or -CSVFilePath"
}

$ErrorActionPreference = "Stop"

function Get-SiteInfo
{
	param([System.String] $SiteUrl)

	try {
		Connect-PnPOnline -Url $SiteUrl -UseWebLogin
		$web = Get-PnPWeb -Includes WebTemplate, Lists, SiteLogoUrl, Language, ParentWeb, RegionalSettings, Configuration
		$site = Get-PnPSite -Includes Owner, Usage
		#$subwebs = Get-PnPSubWebs -Recurse
		$theme = Get-PnPTheme
		$feature = Get-PnPFeature -Scope Site -Identity "7c637b23-06c4-472d-9a9a-7c175762c5c4"
		if ($feature.DefinitionId -eq $null) {
			$lockdownMode = "Disabled"
		} else {
			$lockdownMode = "Enabled"
		}

		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Title" -Value $web.Title
		$obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $web.Url
		$obj | Add-Member -MemberType NoteProperty -Name "Template" -Value "$($web.WebTemplate)#$($web.Configuration)"
		$obj | Add-Member -MemberType NoteProperty -Name "Number of Lists/Libraries" -Value $web.Lists.Count
		#$obj | Add-Member -MemberType NoteProperty -Name "Number of Subwebs" -Value $subwebs.Count
		$obj | Add-Member -MemberType NoteProperty -Name "Limited-access user permission lockdown mode" -Value $lockdownMode
		$obj | Add-Member -MemberType NoteProperty -Name "Site Logo Path" -Value $web.SiteLogoUrl
		$obj | Add-Member -MemberType NoteProperty -Name "Language" -Value $web.Language
		$obj | Add-Member -MemberType NoteProperty -Name "Locale" -Value $web.RegionalSettings.LocaleId
		$obj | Add-Member -MemberType NoteProperty -Name "Theme" -Value $theme.Name
		$obj | Add-Member -MemberType NoteProperty -Name "Owner" -Value $site.Owner.Title
		$obj | Add-Member -MemberType NoteProperty -Name "Size (% Used)" -Value $site.Usage.StoragePercentageUsed

		$obj
	} catch {
		$exceptionName = $_.Exception.GetType().Name
		Write-Host -ForegroundColor "red" "[${exceptionName}] Cannot access: ${SiteUrl}"
	} finally {
		Disconnect-PnPOnline
	}
}

#region main
if ($SiteUrl) {
	Write-Host "Getting Site Info for ${SiteUrl}"
	Get-SiteInfo($SiteUrl)
} else {
	$results = @()
	$csv = Import-Csv $CSVFilePath

	$csv."CO SPO URL" | ForEach-Object {
		Write-Host "Getting Site Info for ${_}"
		$results += Get-SiteInfo($_)
	}

	$results
}
#endregion
