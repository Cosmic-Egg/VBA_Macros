# Define the source and destination directories
$sourceDirectory = "C:\path\to\source"
$destinationDirectory = "C:\path\to\destination"

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