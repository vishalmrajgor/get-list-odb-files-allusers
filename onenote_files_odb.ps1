Connect-MgGraph -ClientId $Env:ENTRA_ID -CertificateThumbprint $Env:Thumbprint -TenantId $Env:Tenant_ID -NoWelcome

Connect-PnPOnline -TenantAdminUrl https://nrdc1-admin.sharepoint.com -ClientId $Env:ENTRA_ID -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint -url https://nrdc1-admin.sharepoint.com
Clear-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
Clear-Content "E:\Automation\Logs\OneDrive_Items_Onenote_US.txt"
Clear-Content "C:\Temp\OneDrive_OneNote_Files.csv"

$Sites = Get-PnPTenantSite -IncludeOneDriveSites | where {$_.url -like "*-my.*" -and $_.Template -notlike "*HOST*"}
$SiteURLs = $Sites.URL
$SiteURLs | Out-File "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
#$Content = Get-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$Output = @()
Foreach($Site in $Sites)
{
$SiteURL = $Site.url
Connect-PnPOnline -ClientId $Env:ENTRA_ID -Url $SiteURL -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint

$Lists = Get-PnPList | where {$_.Hidden -eq $false} | Where {$_.BaseType -eq "DocumentLibrary"} | where {$_.Title -ne "Style Library" -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Site Pages" -and $_.Title -ne "Pages"}   


Foreach($List in $Lists)
{
$List
  try
{
	
	$ListItems = Get-PnPListItem -List $List -PageSize 5000
	$TotalFiles = $ListItems.Count
	$TotalFiles
   	$Count = 0
        
foreach ($Item in $ListItems)
        {
	    #$item
            $FileRef  = $Item["FileRef"]
            $FileLeaf = $Item["FileLeafRef"]
		
	
	    # CONDITION 1: Detect modern OneNote Notebooks (Folders containing a package facet type of 'oneNote')
            $IsOneNotePackage = $null -ne $Item.Package -and $Item.Package.Type -eq "oneNote"
                    
            # CONDITION 2: Detect individual loose section files (.one) or section groups
            $IsOneNoteExtension = $Item.Name -like "*.one" -or $Item.Name -like "*.onetoc2"

		
            if ($FileLeaf -like "*.one" -or $FileLeaf -like "*.onetoc2" -or $IsOneNotePackage -or $IsOneNoteExtension)
            {
		#$Item
                $Count++
		$FileRef
		$FileLeaf
		
		
		$Item.FieldValues.FileLeafRef.ToString()+"|"+$Item.FieldValues.FileRef.ToString() | Out-File "E:\Automation\Logs\OneDrive_Items_Onenote_US.txt" -append
		$Count

                $Output += [PSCustomObject]@{
                    OneDriveUrl = $Site.Url
                    Owner       = $Site.Owner
                    FileName    = $FileLeaf
                    FilePath    = $FileRef
                    Modified    = $Item["Modified"]
                    Created     = $Item["Created"]
		        }
		$FileRef
            }
        }
    }
    catch
    {
        Write-Warning "Failed: $($Site.Url)"
        Write-Warning $_.Exception.Message
    }
}

}
$Output | Export-Csv "C:\Temp\OneDrive_OneNote_Files.csv" -NoTypeInformation

Write-Host "Completed"
