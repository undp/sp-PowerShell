param (
	[Parameter(Mandatory = $true)][string]$siteurl = $(throw "Please specify the URL of the site where you want to create this list"),
	[Parameter(Mandatory = $true)][string]$listTitle = $(throw "Please give the Display Name of the list you wish to create"),
	[Parameter(Mandatory = $true)][string]$listUrl = $(throw "Please specify the URL part of the list - there should be NO SPACES"),
	[Parameter(Mandatory = $true)][string]$csvfilepath = $(throw "Please specify the CSV file of the source list definition")
)

$csvfile = Import-CSV $csvfilepath -Delimiter ","
$counter = 1
$connection = Connect-PnPOnline -Url $siteurl -Interactive
$site = Get-PnPSite
$list = New-PnPList -Title $listTitle -Url $listUrl -Template GenericList


foreach ($filerow in $csvfile) {
	try {
		switch ($filerow.Type) {
			"Choice" {
				$choices = $filerow.Choices.split(",")
				$addfield = Add-PnPField -List $list -Type $filerow.Type -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -Choices $choices -AddToDefaultView
				Break;
			}
			Default {
				$addfield = Add-PnPField -List $list -Type $filerow.Type -DisplayName $filerow.DisplayFieldName -InternalName $filerow.InternalFieldName -AddToDefaultView
			}
		}
		$getfield = Get-PnPField -List $list -Identity $filerow.InternalFieldName
		$getfield.Description = $filerow.Description
		$getfield.Update()
		Write-Host "Added list field/column:" $filerow.InternalFieldName "from CSV row" $counter
		$counter++
	}
	catch {
		Write-host -ForegroundColor "red" -BackgroundColor "black" "Error in the function"
		Write-Host $_.Exception.Message $filerow.InternalFieldName "- at row: " $counter
	}
	finally {
	}
}
