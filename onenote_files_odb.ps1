Connect-MgGraph -ClientId $Env:ENTRA_ID -CertificateThumbprint $Env:Thumbprint -TenantId $Env:Tenant_ID -NoWelcome

Connect-PnPOnline -TenantAdminUrl https://nrdc1-admin.sharepoint.com -ClientId $Env:ENTRA_ID -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint -url https://nrdc1-admin.sharepoint.com
Clear-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$Sites = Get-PnPTenantSite -IncludeOneDriveSites | where {$_.url -like "*-my.*" -and $_.Template -notlike "*HOST*"}
$SiteURLs = $Sites.URL
$SiteURLs | Out-File "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$Content = Get-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"

Foreach($SiteURL in $Content)
{
Connect-PnPOnline -ClientId $Env:ENTRA_ID -Url $SiteURL -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint

$Lists = Get-PnPList | where {$_.Hidden -eq $false} | Where {$_.BaseType -eq "DocumentLibrary"} | where {$_.Title -ne "Style Library" -and $_.Title -ne "Form Templates" -and $_.Title -ne "Site Assets" -and $_.Title -ne "Site Pages" -and $_.Title -ne "Pages"}   



Foreach($List in $Lists)
{

  try
{

    $ListItems = Get-PnPListItem -List "Documents" -PageSize 5000
	$TotalFiles = $ListItems.Count
    $Count = 0
        foreach ($Item in $ListItems)
        {
            $FileRef  = $Item["FileRef"]
            $FileLeaf = $Item["FileLeafRef"]

	        # CONDITION 1: Detect modern OneNote Notebooks (Folders containing a package facet type of 'oneNote')
            $IsOneNotePackage = $null -ne $Item.Package -and $Item.Package.Type -eq "oneNote"
                    
            # CONDITION 2: Detect individual loose section files (.one) or section groups
            $IsOneNoteExtension = $Item.Name -like "*.one" -or $Item.Name -like "*.onetoc2"

		
            if ($FileLeaf -like "*.one" -or $FileLeaf -like "*.onetoc2" -or $IsOneNotePackage -or $IsOneNoteExtension)
            {
                $Count++
		        $ListItem.FieldValues.FileLeafRef.ToString()+"|"+$ListItem.FieldValues.FileRef.ToString()+"|"+$TodaysDate.ToString()+"|"+$TimeNow.ToString() | Out-File "E:\Automation\Logs\OneDrive_Items_Onenote_US.txt" -append
                $PercentComplete = (($Count / $TotalFiles) * 100)
                $CurrentOperation = "Processing Item $Count of $TotalFiles"
                $Activity = "marking onenote files"
                $PSStyle.Progress.View = "Classic"
                Write-Progress -PercentComplete $PercentComplete -CurrentOperation $CurrentOperation -Activity $Activity

                $Output += [PSCustomObject]@{
                    OneDriveUrl = $SiteURL.Url
                    Owner       = $SiteURL.Owner
                    FileName    = $FileLeaf
                    FilePath    = $FileRef
                    Modified    = $Item["Modified"]
                    Created     = $Item["Created"]
		        }
            }
        }
    }
    catch
    {
        Write-Warning "Failed: $($SiteURL.Url)"
        Write-Warning $_.Exception.Message
    }
}

}
$Output | Export-Csv "C:\Temp\OneDrive_OneNote_Files.csv" -NoTypeInformation

Write-Host "Completed"
