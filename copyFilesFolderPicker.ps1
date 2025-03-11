Add-Type -AssemblyName "System.Windows.Forms"

# Define the fixed destination folder
$destinationFolder = "C:\Path\To\Destination"

# Create and show the folder picker dialog
$folderDialog = New-Object Windows.Forms.FolderBrowserDialog
$folderDialog.Description = "Select the source folder"
$folderDialog.ShowNewFolderButton = $false

# Show the dialog and get the selected folder path
$dialogResult = $folderDialog.ShowDialog()

if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
    $sourceFolder = $folderDialog.SelectedPath

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
    Write-Host "No folder selected. Exiting script."
}
