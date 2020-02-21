<#
	.SYNOPSIS
	PowerShell cmdlet to set the logo of a web, and optionally, all subwebs of a site collection

	.DESCRIPTION
	PowerShell cmdlet that sets the sitelogo of a site collection, and optionally, all the subwebs in that site collection

	.PARAMETER webUrl
	Specifies the webUrl where the site logo is to be updated

	.PARAMETER recursion
	Optional parameter, specifies whether the site logo is to be applied on all of the subwebs. Default is false

	.PARAMETER siteLogoPath
	Specifies the path of the site logo that should be applied. This string should only contain the path, not the full path which includes the host

	.INPUTS
	The Microsoft .NET Framework types of objects that can be piped to the function or script. You can also include a description of the input objects.
	None. You cannot pipe objects to Add-Extension.

	.OUTPUTS
	The .NET Framework type of the objects that the cmdlet returns. You can also include a description of the returned objects.
	System.String. Add-Extension returns a string with the extension
	or file name.

	.EXAMPLE
	A sample command that uses the function or script, optionally followed by sample output and a description. Repeat this keyword for each example.
	C:\PS> Set-SPOSiteLogo.ps1 -webUrl "https://tenant.sharepoint.com/teams/ABC" -recursion $true -siteLogoPath "/SiteAssets/sitelogo.png"
	
	.EXAMPLE
	C:\PS> Set-SPOSiteLogo.ps1 -webUrl "https://tenant.sharepoint.com/teams/ABC" -siteLogoPath "/SiteAssets/sitelogo.png"
	File.doc

	.EXAMPLE
	C:\PS> extension "File" "doc"
	File.doc

	.LINK
	Script inspired by Jason Lee: http://www.jrjlee.com/2016/12/applying-logo-to-every-site-in.html

	.LINK
	Set-Item

	.NOTES
AUTHOR: Daniel Tshin
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 24 May 2019
#
# version: 1.1
#
# Revision history
# # 24 May 2019 Creation
# # 06 Jun 2019 Added Recursive function, re-factored script

#>


#region command line parameters ###############################################
# Input Parameters
# $args: URL of the web hosting the list or library, the name of the list or library to inspect. Separated by space
param(
    [Parameter(Mandatory=$true)][string]$webUrl = $(throw "Please specify the URL for the site where you want to change the site logo"),
	[Parameter(Mandatory=$false)][string]$recursion,
    [Parameter(Mandatory=$true)][string]$siteLogoPath = $(throw "Please specify the path of the site logo. it should start with '/' and should not include the host ('https://tenant.sharepoint.com/')")
	);
#endregion ####################################################################
Function update-webLogo {
	Param(
		[parameter(Mandatory=$true)][string]$serverRelativeUrl
	)
	Begin { }
	Process {
		$web = Get-PnPWeb -identity $serverRelativeUrl
		Write-OutPut ("Setting site logo on subweb: {0}" -f $web.serverRelativeUrl)
		Set-PnPWeb -Web $serverRelativeUrl -siteLogoUrl $siteLogoPath
	}
}

Function Set-SPOSiteLogo {
	Process {
		Try {
			# For SPO
			Connect-PnPOnline -Url $webUrl -UseWebLogin
			# End SPO
			
			$rootWeb = Get-PnPWeb

			#Set logo on root web
			Set-PnPWeb -Web $rootWeb -siteLogoUrl $siteLogoPath
			Write-OutPut ("Setting site logo on web: {0}" -f $rootWeb.serverRelativeUrl)

			if ($recursion -eq $true) {
				$subwebs = Get-PnPSubWebs -Web $rootWeb -Recurse
				foreach ($subweb in $subwebs) {
					update-webLogo($subweb.serverRelativeUrl)
				}
			}

		}
		Catch {
			write-host -f Red "Error in Script: " $_.Exception.Message
		}
		Finally {
		}
	}
}

#region [ Main Execution ]----------------------------------------------------

# Script Execution goes here
Set-SPOSiteLogo
#endregion 