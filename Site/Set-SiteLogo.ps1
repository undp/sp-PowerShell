<#
	.SYNOPSIS
		Set the logo of a web, and optionally, all subwebs of a site collection

	.DESCRIPTION
		The Set-SiteLogo script sets the site logo of a site collection, and optionally, all the subwebs in that site collection.

	.PARAMETER WebUrl
		Specifies the WebUrl where the site logo is to be updated.

	.PARAMETER Recursion
		Optional parameter, specifies whether the site logo is to be applied on all of the subwebs. Default is false.

	.PARAMETER SiteLogoPath
		Specifies the path of the site logo that should be applied. This string should only contain the path, not the full path which includes the host.

	.EXAMPLE
		Set-SiteLogo.ps1 -WebUrl "https://tenant.sharepoint.com/teams/MyTeam" -Recursion $true -SiteLogoPath "/SiteAssets/SiteLogo.png"
	
	.EXAMPLE
		Set-SiteLogo.ps1 -WebUrl "https://tenant.sharepoint.com/teams/MyTeam" -SiteLogoPath "/SiteAssets/SiteLogo.png"

	.LINK
		Script inspired by Jason Lee: http://www.jrjlee.com/2016/12/applying-logo-to-every-site-in.html

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
#>

param(
    [Parameter(Mandatory=$true)][string]$WebUrl = $(throw "Please specify the URL for the site where you want to change the site logo."),
	[Parameter(Mandatory=$false)][string]$Recursion,
    [Parameter(Mandatory=$true)][string]$SiteLogoPath = $(throw "Please specify the path of the site logo. it should start with '/' and should not include the host ('https://tenant.sharepoint.com/').")
)

function Set-WebLogo {
	param(
		[parameter(Mandatory=$true)][string]$ServerRelativeUrl
	)
	
	$web = Get-PnPWeb -Identity $ServerRelativeUrl
	Write-Output ("Setting site logo on subweb: {0}" -f $web.ServerRelativeUrl)
	Set-PnPWeb -Web $ServerRelativeUrl -SiteLogoUrl $SiteLogoPath
}

#region main
try {
	Connect-PnPOnline -Url $WebUrl -UseWebLogin
	$rootWeb = Get-PnPWeb

	# Set logo on root web
	Set-PnPWeb -Web $rootWeb -siteLogoUrl $SiteLogoPath
	Write-OutPut ("Setting site logo on web: {0}" -f $rootWeb.ServerRelativeUrl)

	if ($Recursion -eq $true) {
		$subwebs = Get-PnPSubWebs -Web $rootWeb -Recurse
		foreach ($subweb in $subwebs) {
			Set-WebLogo($subweb.ServerRelativeUrl)
		}
	}

} catch {
	Write-Host -f Red "Error in Script: " $_.Exception.Message
} finally {
	Disconnect-PnPOnline
}
#endregion
