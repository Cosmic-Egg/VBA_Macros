# Define the fixed destination folder
$destinationFolder = "C:\Path\To\Destination"

# Get a list of subdirectories (folders) in the source parent directory
$sourceParentFolder = "C:\Path\To\Source"
$availableFolders = Get-ChildItem -Path $sourceParentFolder -Directory

# Display the available folders for the user to choose from with numbers
Write-Host "Please select a source folder from the list below:"
for ($i = 0; $i -lt $availableFolders.Count; $i++) {
    Write-Host "$($i + 1). $($availableFolders[$i].Name)"
}

# Prompt the user to choose a folder by number
$folderChoice = Read-Host "Enter the number of the folder you want to select"

# Check if the input is valid
if ($folderChoice -gt 0 -and $folderChoice -le $availableFolders.Count) {
    # Get the selected source folder
    $selectedFolder = $availableFolders[$folderChoice - 1]

    # Define the full source folder path
    $sourceFolder = $selectedFolder.FullName

    # Check if the destination folder exists, if not, create it
    if (-Not (Test-Path -Path $destinationFolder)) {
        Write-Host "Destination folder does not exist. Creating folder..."
        New-Item -Path $destinationFolder -ItemType Directory
    }

    # Get the list of files from the selected source folder
    $files = Get-ChildItem -Path $sourceFolder

    # Copy each file to the fixed destination folder
    foreach ($file in $files) {
        $destinationFile = Join-Path -Path $destinationFolder -ChildPath $file.Name
        Write-Host "Copying $($file.Name) to $destinationFolder"
        Copy-Item -Path $file.FullName -Destination $destinationFile
    }

    Write-Host "Files copied successfully from $sourceFolder to $destinationFolder."
} else {
    Write-Host "Invalid choice. Please restart the script and select a valid source folder number."
}
