<#
	.SYNOPSIS
		Get-SPOSiteInfo gets information about SharePoint Online Sites

	.DESCRIPTION
		Get-SPOSiteInfo gets information of SharePoint Online sites including Title, URL, WebTemplate, List Count, Subwebs Count, MasterPageURL, InheritsMasterURL, InheritsCustomMasterURL, Owner, Size, Language, Locale, ParentWebId, SiteLogoPath. Get-SPOSiteInfo uses SharePoint-PnPPowerShell
Takes a string site url or  csv file path.
Information is output as standard PSObject.

	.PARAMETER siteURL
		Root Site URL of the webs to inspect.
	
	.PARAMETER csvFilePath
		CSV File path containing Site URLs of the webs to inspect.

	.EXAMPLE
		Get-SPOSiteInfo -siteUrl https://tenant.sharepoint.com/Sites/CollectionURL

	.EXAMPLE
		Get-SPOSiteInfo -csvFilePath "C:/csvfolder/mycsvfile.csv"

	.INPUTS
		System.String,System.Int32

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
		AUTHOR: Giuseppe Campanelli, giuseppe.campanelli@undp.org
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 04 October 2019
# Modified by: Giuseppe Campanelli (giuseppe.campanelli@undp.org) on 31 March 2020
#
# version: 1.1
#
# Revision history
# # 11 Jun 2012 Creation
# # 22 Jun 2012 Gets list of site collection URLs from command line $args, also checks and outputs if UNDPCOInfoModule feature is activated
# # 22 Jun 2012 Script checks for site collection level and does not process, change name to "Get-WebInfo" to correctly reflect its purpose
# # 06 Jul 2012 Reorganizing the script, re-wrote Log-Message function
# # 07 Sept 2012 Added queries for properties: InheritsAlternateCssURL and AlternateCSSUrl
# # 06 Feb 2013 Added Finally clause with web.dispose()
# # 05 Feb 2014 Added query for property: number of lists
# # 04 Oct 2019 Forked new cmdlet: "Get-SPOSiteInfo" for getting details of SPO sites via SharePoint-PnPPowerShell
# # 31 Mar 2020 Revised script to use PnP, accept csv with site urls and update error handling

.FUNCTIONALITY
		Get-SPOSiteInfo is to be used to get details of webs of a site url or from site urls specified in a csv file. Output can be processed or exported.

	.LINK
		about_functions_advanced

	.LINK
		about_comment_based_help

#>


#region command line parameters ###############################################
# Input Parameters
# $siteUrl: url of the site to get info of
# $csvFilePath: csv file path containing site urls
param(
	[Parameter(Mandatory=$false)][string]$siteUrl,
	[Parameter(Mandatory=$false)][string]$csvFilePath
	);

if ($psboundparameters.Count -ne 1) {
	Write-Host "Pleasy specify only one of three parameters: -siteURl or -csvFilePath"
}
#endregion ####################################################################

#region script settings #######################################################
$ErrorActionPreference = "Stop"
#endregion ####################################################################

###################################
## helper functions
###################################

#region GetSPOInfo ############################################################
###############################################################################
## Function GetSPOInfo
## Returns web info of SPO site.
##
## Paramters
# # $siteUrl			- string: url of the site
###############################################################################
function GetSPOInfo
{
	param([System.String] $siteUrl)

	try {
		Connect-PnPOnline -Url $siteUrl -UseWebLogin
		$web = Get-PnPWeb -Includes WebTemplate, Lists, MasterUrl, AllProperties, AlternateCSSUrl, SiteLogoUrl, Language, ParentWeb, RegionalSettings
		$site = Get-PnPSite -Includes Owner, Usage
		#$subwebs = Get-PnPSubWebs -Recurse
		$feature = Get-PnPFeature -Scope Site -Identity "7c637b23-06c4-472d-9a9a-7c175762c5c4"
		Write-Host $feature.Count
		if ($feature.DefinitionId -eq $null) {
			$lockdownMode = "Disabled"
		} else {
			$lockdownMode = "Enabled"
		}

		$obj = New-Object PSObject
		$obj | Add-Member -MemberType NoteProperty -Name "Title" -Value $web.Title
		$obj | Add-Member -MemberType NoteProperty -Name "URL" -Value $web.Url
		$obj | Add-Member -MemberType NoteProperty -Name "Template" -Value $web.WebTemplate
		$obj | Add-Member -MemberType NoteProperty -Name "Number of Lists/Libraries" -Value $web.Lists.Count
		#$obj | Add-Member -MemberType NoteProperty -Name "Number of Subwebs" -Value $subwebs.Count
		$obj | Add-Member -MemberType NoteProperty -Name "Limited-access user permission lockdown mode" -Value $lockdownMode
		$obj | Add-Member -MemberType NoteProperty -Name "Site Logo Path" -Value $web.SiteLogoUrl
		$obj | Add-Member -MemberType NoteProperty -Name "Language" -Value $web.Language
		$obj | Add-Member -MemberType NoteProperty -Name "Owner" -Value $site.Owner.Title
		$obj | Add-Member -MemberType NoteProperty -Name "Size (% Used)" -Value $site.Usage.StoragePercentageUsed
		$obj | Add-Member -MemberType NoteProperty -Name "Parent Web Id" -Value $web.ParentWeb.Id
		$obj | Add-Member -MemberType NoteProperty -Name "Locale" -Value $web.RegionalSettings.LocaleId
		$obj | Add-Member -MemberType NoteProperty -Name "MasterPageURL" -Value $web.MasterUrl
		$obj | Add-Member -MemberType NoteProperty -Name "InheritsMasterURL" -Value $web.AllProperties["__InheritsMasterUrl"]
		$obj | Add-Member -MemberType NoteProperty -Name "InheritsCustomMasterURL" -Value $web.AllProperties["__InheritsCustomMasterUrl"]
		$obj | Add-Member -MemberType NoteProperty -Name "InheritsAlternateCssURL" -Value $web.AllProperties["__InheritsAlternateCssUrl"]
		$obj | Add-Member -MemberType NoteProperty -Name "AlternateCssURL" -Value $web.AlternateCssUrl

		$obj
	} catch {
		$exceptionName = $_.Exception.GetType().Name
		Write-Host -ForegroundColor "red" "[${exceptionName}] Cannot access: ${siteUrl}"
	} finally {
		Disconnect-PnPOnline
	}
}
#endregion ################################################################

#region MAIN ##############################################################
########################## MAIN ###########################################
if ($siteUrl) {
	Write-Host "Getting SPO Site Info for ${siteUrl}"
	GetSPOInfo($siteUrl)
} else {
	$results = @()
	$csv = Import-Csv $csvFilePath

	$csv."CO SPO URL" | ForEach-Object {
		Write-Host "Getting SPO Site Info for ${_}"
		$results += GetSPOInfo($_)
	}

	$results
}
#endregion ################################################################