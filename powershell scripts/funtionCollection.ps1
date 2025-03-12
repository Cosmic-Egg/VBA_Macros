function Copy-Files {
    param (
        [string]$destinationFolder = "C:\Path\To\Destination",
        [string]$sourceParentFolder = "C:\Path\To\Source"
    )

    # Get a list of subdirectories (folders) in the source parent directory
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
}
function Copy-FilesFolderPicker {
    param (
        [string]$destinationFolder = "C:\Path\To\Destination"
    )

    Add-Type -AssemblyName "System.Windows.Forms"

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
}

function Copy-MostRecentFiles {
    param (
        [string]$parentFolder = "C:\Path\To\ParentFolder",
        [string]$destinationFolder = "C:\Path\To\DestinationFolder"
    )

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
}

function Add-DefaultFolders {
    param (
        [string]$sourceDirectory = "C:\path\to\source",
        [string]$destinationDirectory = "C:\path\to\destination"
    )

    # Define the folder names you want to move
    $folderNames = @(
        "Archive",
        "Assumptions",
        "Final Deliverables",
        "Model and IT",
        "Testing and Analysis"
    )

    # Move each folder
    foreach ($folderName in $folderNames) {
        $sourcePath = Join-Path -Path $sourceDirectory -ChildPath $folderName
        $destinationPath = Join-Path -Path $destinationDirectory -ChildPath $folderName
        Move-Item -Path $sourcePath -Destination $destinationPath -Force
    }
}

