<#
	.SYNOPSIS
		Initialize-SPOSite Initializes and performs configuration for SharePoint Online Site Collection(s)

	.DESCRIPTION
		Initialize-SPOSite Initializes and performs basic configuration for SharePoint Online Site Collection(s). Unit owner is added to Site Collection Administrator and the Owners group. 

	.PARAMETER
		Root Site Collection URL of the SPO site to initailize and configure. 

	.EXAMPLE
		Initialize-SPOSite -siteUrl https://undp.sharepoint.com/teams/IND
	
	.EXAMPLE
		Initialize-SPOSite -siteUrl https://undp.sharepoint.com/teams/IND -SCAdmin user.name@undp.org

	.INPUTS
		System.String,System.Int32

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR: Daniel Tshin
LASTEDIT: $(Get-Date)
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 28 April 2017
#
# version: 1.0
#
# Revision history
# # 28 April 2017 Creation

.FUNCTIONALITY
		Initialize-SPOSite is to be used to initialize configuration of SPO site collections. Unit owner is added to Site Collection Administrator and the Owners group. Output can be processed or exported. 
		
	.LINK
		about_functions_advanced

	.LINK
		about_comment_based_help

#>

param([string]$siteUrl = $(Read-Host -Prompt "Input the URL of the SPO site collection"))
#region command line parameters ###############################################
# Input Parameters
# $siteUrl: Base URL of SharePoint Online site collection 
# $siteUrl = "https://undp.sharepoint.com/teams/IND"  
#endregion ####################################################################

#region #######################################################################
# PowerShell script for getting group (user) information from SharePoint Online webs
# Written by: Daniel Tshin (daniel.tshin@undp.org) on 28 April 2017
#
# version: 1.0
#
# Revision history
# # 28 April 2017 Creation
#
#endregion ####################################################################

#region Setup constants #######################################################
# Sets full path to CSVFilename
$LogFile ="$pwd\InitializeSPOSiteLogFile.txt"
$ErrorLogFile ="$pwd\InitializeSPOSiteErrorLogFile.txt"
## SharePoint DLL
try {
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
	Write-Host "Loading Assemblies`n" -ForegroundColor Magenta
	$SPOMgmtShellAssembyPath = resolve-path("C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell") -ErrorAction Stop
	#$TenantAssembyPath = resolve-path("C:\Program Files\SharePoint Client Components\16.0\Assemblies") -ErrorAction Stop
	Add-Type -Path ($SPOMgmtShellAssembyPath + "\Microsoft.SharePoint.Client.dll")
	Add-Type -Path ($SPOMgmtShellAssembyPath + "\Microsoft.SharePoint.Client.Runtime.dll"  

} catch {
	Write-Host "Can't load assemblies..." -ForegroundColor Red
	Write-Host $Error[0].Exception.Message -ForegroundColor Red
    exit
}
#Add-PSSnapin Microsoft.SharePoint.PowerShell
#
#endregion ####################################################################


###################################

#region Log-Message #### ######################################################
###################################################################################
## Function Log-Message
# # Logs messages to custom log file
##
## Paramters
# # $message	- message - to be logged
# # $weburl		- url: web URL: web
# # $LogFile	- string: logfile (global)
###################################################################################
function Log-Message($message, $weburl)
{   
	Write-Output ($message, $weburl)| Out-File -FilePath $LogFile -Append 
}
#endregion ####################################################################



#$userId = "daniel.tshin@undp.org"
$userId = Read-Host -Prompt "Enter your UserID (email address)"
#$pwd = Read-Host -Prompt "Enter password" -AsSecureString  
#$creds = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userId, $pwd)  
$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($siteUrl)  
#$ctx.credentials = $creds  
$ctx.credentials = Get-Credentials  -UserName $userId -message "Enter your O365 credentials"
try{  
    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle("TestList")  
    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $ctx.load($listItems)  
      
    $ctx.executeQuery()  
    foreach($listItem in $listItems)  
    {  
        Write-Host "ID - " $listItem["ID"] "Title - " $listItem["Title"]  
    }  
}  
catch{  
    write-host "$($_.Exception.Message)" -foregroundcolor red  
}  