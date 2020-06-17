<#
	.SYNOPSIS
		Initialize and perform configuration for SharePoint Online Site Collection.

	.DESCRIPTION
		The Init-SPOSite initializes and performs basic configuration for SharePoint Online Site Collection(s). Unit owner is added to Site Collection Administrator and the Owners group. 

	.PARAMETER SiteUrl
		Root Site Collection URL of the SPO site to initailize and configure. 

	.EXAMPLE
		Init-SPOSite.ps1 -SiteUrl "https://tenant.sharepoint.com/teams/MyTeam"

	.OUTPUTS
		PSObject

	.NOTES
		AUTHOR: Daniel Tshin, daniel.tshin@undp.org
#>

param(
	[string]$SiteUrl = $(Read-Host -Prompt "Input the URL of the SPO site collection.")
)

#region Setup constants
# SharePoint DLL
try {
	[void][System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint") | Out-Null
	Write-Host "Loading Assemblies`n" -ForegroundColor Magenta
	$SPOMgmtShellAssembyPath = Resolve-Path("C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell") -ErrorAction Stop
	#$TenantAssembyPath = Resolve-Path("C:\Program Files\SharePoint Client Components\16.0\Assemblies") -ErrorAction Stop
	Add-Type -Path ($SPOMgmtShellAssembyPath + "\Microsoft.SharePoint.Client.dll")
	Add-Type -Path ($SPOMgmtShellAssembyPath + "\Microsoft.SharePoint.Client.Runtime.dll"  

} catch {
	Write-Host "Can't load assemblies..." -ForegroundColor Red
	Write-Host $Error[0].Exception.Message -ForegroundColor Red
    exit
}
#Add-PSSnapin Microsoft.SharePoint.PowerShell
#endregion

#region main
try {
	$userId = Read-Host -Prompt "Enter your UserID (email address)."
	$ctx = New-Object Microsoft.SharePoint.Client.ClientContext($SiteUrl)  
	$ctx.Credentials = Get-Credentials -UserName $userId -message "Enter your O365 credentials."

    $lists = $ctx.web.Lists  
    $list = $lists.GetByTitle("TestList")  
    $listItems = $list.GetItems([Microsoft.SharePoint.Client.CamlQuery]::CreateAllItemsQuery())  
    $ctx.Load($listItems)  
      
    $ctx.ExecuteQuery()  
    foreach($listItem in $listItems)  
    {  
        Write-Host "ID - " $listItem["ID"] "Title - " $listItem["Title"]  
    }  
}  catch {  
    Write-Host "$($_.Exception.Message)" -foregroundcolor red  
}
#endregion
