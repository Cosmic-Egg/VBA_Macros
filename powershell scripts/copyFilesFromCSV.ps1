Add-Type -AssemblyName "System.Windows.Forms"

function Show-FolderPicker {
    param (
        [string]$description = "Select a folder"
    )
    $folderDialog = New-Object Windows.Forms.FolderBrowserDialog
    $folderDialog.Description = $description
    $folderDialog.ShowNewFolderButton = $true
    $dialogResult = $folderDialog.ShowDialog()
    if ($dialogResult -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderDialog.SelectedPath
    } else {
        Write-Host "No folder selected. Exiting script."
        exit
    }
}

function Search-FilesByPartialName {
    param (
        [string]$csvFilePath = "C:\Path\To\PartialNames.csv"
    )

    # Use folder picker for source and destination folders
    $searchDirectory = Show-FolderPicker -description "Select the search directory"
    $destinationDirectory = Show-FolderPicker -description "Select the destination directory"

    # Import the CSV file
    $partialNames = Import-Csv -Path $csvFilePath

    # Check if the destination directory exists, if not, create it
    if (-Not (Test-Path -Path $destinationDirectory)) {
        Write-Host "Destination directory does not exist. Creating directory..."
        New-Item -Path $destinationDirectory -ItemType Directory
    }

    # Search for files with partial names and copy them to the destination directory
    foreach ($partialName in $partialNames) {
        $files = Get-ChildItem -Path $searchDirectory -Recurse -File | Where-Object { $_.Name -like "*$($partialName.PartialName)*" }
        foreach ($file in $files) {
            $destinationFile = Join-Path -Path $destinationDirectory -ChildPath $file.Name
            Write-Host "Copying '$($file.FullName)' to '$destinationFile'"
            Copy-Item -Path $file.FullName -Destination $destinationFile -Force
        }
    }

    Write-Host "Files with partial names from the CSV have been copied to $destinationDirectory."
}

# Example usage
Search-FilesByPartialName -csvFilePath "C:\Path\To\PartialNames.csv"