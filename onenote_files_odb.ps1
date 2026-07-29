Connect-MgGraph -ClientId $Env:ENTRA_ID -CertificateThumbprint $Env:Thumbprint -TenantId $Env:Tenant_ID -NoWelcome
Connect-PnPOnline -TenantAdminUrl https://nrdc1-admin.sharepoint.com -ClientId $Env:ENTRA_ID -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint -url https://nrdc1-admin.sharepoint.com
Clear-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$Sites = Get-PnPTenantSite -IncludeOneDriveSites | where {$_.url -like "*-my.*" -and $_.Template -notlike "*HOST*"}
$SiteURLs = $Sites.URL
$SiteURLs | Out-File "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$Content = Get-Content "E:\Automation\Scripts\Temp\Onenote_All_OneDrive_US.txt"
$ListName= "Documents"
$ViewName= "All"
Foreach($SiteURL in $Content)
{
Connect-PnPOnline -ClientId $Env:ENTRA_ID -Url $SiteURL -Tenant nrdc1.onmicrosoft.com -Thumbprint $Env:Thumbprint

    try
    {

        $Items = Get-PnPListItem -List "Documents" -PageSize 5000

        foreach ($Item in $Items)
        {
            $FileRef  = $Item["FileRef"]
            $FileLeaf = $Item["FileLeafRef"]

            if ($FileLeaf -like "*.one")
            {
                $Output += [PSCustomObject]@{
                    OneDriveUrl = $OD.Url
                    Owner       = $OD.Owner
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
        Write-Warning "Failed: $($OD.Url)"
        Write-Warning $_.Exception.Message
    }
}

$Output | Export-Csv "C:\Temp\OneDrive_OneNote_Files.csv" -NoTypeInformation

Write-Host "Completed"
