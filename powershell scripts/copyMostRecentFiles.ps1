# Define the parent folder (where your subfolders are located)
$parentFolder = "C:\Path\To\ParentFolder"

# Define the fixed destination folder
$destinationFolder = "C:\Path\To\DestinationFolder"

# Get all subdirectories (folders) in the parent folder
$subFolders = Get-ChildItem -Path $parentFolder -Directory

# Check if the destination folder exists, if not, create it
if (-Not (Test-Path -Path $destinationFolder)) {
    Write-Host "Destination folder does not exist. Creating folder..."
    New-Item -Path $destinationFolder -ItemType Directory
}

# Loop through each subfolder to find the most recent file
foreach ($subFolder in $subFolders) {
    # Get the most recent file in the subfolder
    $mostRecentFile = Get-ChildItem -Path $subFolder.FullName -File | Sort-Object LastWriteTime -Descending | Select-Object -First 1

    # Check if a file was found in the subfolder
    if ($mostRecentFile) {
        # Define the destination file path
        $destinationFile = Join-Path -Path $destinationFolder -ChildPath $mostRecentFile.Name

        # Copy the most recent file to the destination folder
        Write-Host "Copying '$($mostRecentFile.Name)' from '$($subFolder.Name)' to '$destinationFolder'"
        Copy-Item -Path $mostRecentFile.FullName -Destination $destinationFile -Force
    } else {
        Write-Host "No files found in folder '$($subFolder.Name)'."
    }
}

Write-Host "Most recent files from each subfolder have been copied to $destinationFolder."
