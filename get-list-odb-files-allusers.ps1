# Connect with the required scopes for cross-user OneNote auditing
Connect-MgGraph -Scopes "User.Read.All", "Files.Read.All", "Directory.Read.All", "Notes.Read.All"

# Define the output report path
$ReportPath = "E:\Temp\AllUsersOneNoteInventory.csv"
$ReportResults = @()

# 1. Fetch all licensed user members in the organization
Write-Host "Retrieving active user directory..." -ForegroundColor Cyan
$Users = Get-MgUser -Filter "assignedLicenses/`$count ne 0 and userType eq 'Member'" -ConsistencyLevel eventual -CountVariable Records -All

# 2. Iterate through each user to extract OneNote items
foreach ($User in $Users) {
    Write-Host "Searching OneNote items for: $($User.UserPrincipalName)" -ForegroundColor Yellow
    
    try {
        # Retrieve the user's primary default OneDrive drive
        $UserDrive = Get-MgUserDefaultDrive -UserId $User.Id -ErrorAction Stop
        
        if ($UserDrive) {
            # Internal recursive helper function to search through folder structures
            function Get-OneNoteItemsRecursive {
                param (
                    [string]$UserId,
                    [string]$DriveId,
                    [string]$ItemId = "root",
                    [string]$CurrentPath = ""
                )
                
                # Fetch children of the current folder item
                $Items = Get-MgDriveItemChild -DriveId $DriveId -DriveItemId $ItemId -All -ErrorAction SilentlyContinue
                
                foreach ($Item in $Items) {
                    $FullPath = if ($CurrentPath -eq "") { $Item.Name } else { "$CurrentPath/$($Item.Name)" }
                    
                    # CONDITION 1: Detect modern OneNote Notebooks (Folders containing a package facet type of 'oneNote')
                    $IsOneNotePackage = $null -ne $Item.Package -and $Item.Package.Type -eq "oneNote"
                    
                    # CONDITION 2: Detect individual loose section files (.one) or section groups
                    $IsOneNoteExtension = $Item.Name -like "*.one" -or $Item.Name -like "*.onetoc2"
                    
                    if ($IsOneNotePackage -or $IsOneNoteExtension) {
                        # Item found; record the properties into the master data collection array
                        $Global:ReportResults += [PSCustomObject]@{
                            UserPrincipalName = $User.UserPrincipalName
                            NotebookName      = $Item.Name
                            ItemType          = if ($IsOneNotePackage) { "OneNote Notebook" } else { "Loose Section File (.one)" }
                            OneDrivePath      = $FullPath
                            WebUrl            = $Item.WebUrl
                            LastModified      = $Item.LastModifiedDateTime
                        }
                        
                        # Notebook folders don't need further recursion down their internal system subfolders
                        continue
                    }
                    
                    # If it's a regular directory folder, recurse deeper to search for notebooks inside it
                    if ($Item.Folder) {
                        Get-OneNoteItemsRecursive -UserId $UserId -DriveId $DriveId -DriveItemId $Item.Id -CurrentPath $FullPath
                    }
                }
            }
            
            # Execute the recursive analysis beginning at the root level
            Get-OneNoteItemsRecursive -UserId $User.Id -DriveId $UserDrive.Id
        }
    }
    catch {
        Write-Host "Could not access data or no personal drive provisioned for $($User.UserPrincipalName)." -ForegroundColor DarkGray
    }
}

# 3. Export data collection cleanly to a local file
if ($ReportResults.Count -gt 0) {
    New-Item -ItemType Directory -Force -Path (Split-Path $ReportPath) | Out-Null
    $ReportResults | Export-Csv -Path $ReportPath -NoTypeInformation -Encoding UTF8
    Write-Host "Success! The OneNote inventory has been exported to: $ReportPath" -ForegroundColor Green
} else {
    Write-Host "No OneNote notebooks or sections were discovered across user environments." -ForegroundColor DarkYellow
}
