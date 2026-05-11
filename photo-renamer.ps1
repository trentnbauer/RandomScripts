<#
Trent's Photography Script
Partially written by Claude

This script was written to help with sorting my images by year, season and down to camera and film stock.
I will then sort the folder by creation date.

The intention is to create a yearly photobook, sorted by season. I can also add the camera model and film stock as a note in the book if I want.
#>

$SourcePath = $null
$OutputPath = $null
$Camera = $null
$FilmStock = $null
$Hemisphere = $null  # Set to "Northern" or "Southern", or leave $null to be prompted

# Other variables you probably don't need to edit
$WorkingDir = "$env:TEMP\photoscript"
$RegPath = "HKCU:\Software\TrentsPhotoScript"

#### ---- FUNCTIONS ---- ####

function PickFolder ([string]$Description = "Select a folder") {
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = $Description
    $FolderBrowser.RootFolder = "MyComputer"

    if ($FolderBrowser.ShowDialog() -eq "OK") {
        $Selected = $FolderBrowser.SelectedPath
    } else {
        throw "No folder selected."
    }

    if (-not (Test-Path $Selected)) {
        throw "Selected path does not exist: $Selected"
    }
    return $Selected
}

function Test-Writable ([string]$Path) {
    $TestFile = Join-Path $Path "write_test_$((New-Guid).Guid).tmp"
    try {
        New-Item -Path $TestFile -ItemType File -ErrorAction Stop | Out-Null
        Remove-Item $TestFile -ErrorAction Stop | Out-Null
    } catch {
        throw "Path is not writable: $Path`n$($_.Exception.Message)"
    }
}

function Test-IsImage ([string]$FilePath) {
    Add-Type -AssemblyName System.Web
    $MimeType = [System.Web.MimeMapping]::GetMimeMapping($FilePath)
    return $MimeType -like "image/*"
}

function Get-Season ([datetime]$Date) {
    $month = $Date.Month

    if ($Hemisphere -eq "Southern") {
        switch ($month) {
            { $_ -in 9, 10, 11  } { return "Spring" }
            { $_ -in 12, 1, 2   } { return "Summer" }
            { $_ -in 3, 4, 5    } { return "Autumn" }
            { $_ -in 6, 7, 8    } { return "Winter" }
        }
    } else {
        switch ($month) {
            { $_ -in 3, 4, 5    } { return "Spring" }
            { $_ -in 6, 7, 8    } { return "Summer" }
            { $_ -in 9, 10, 11  } { return "Autumn" }
            { $_ -in 12, 1, 2   } { return "Winter" }
        }
    }
}

function Get-SavedSettings {
    if (Test-Path $RegPath) {
        $RegValues = Get-ItemProperty -Path $RegPath -ErrorAction SilentlyContinue
        if ($RegValues.OutputPath -and $RegValues.Hemisphere) {
            return $RegValues
        }
    }
    return $null
}

function Save-Settings ([string]$OutputPath, [string]$Hemisphere) {
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $RegPath -Name "OutputPath" -Value $OutputPath
    Set-ItemProperty -Path $RegPath -Name "Hemisphere" -Value $Hemisphere
}

#### ---- START SCRIPT ---- ####

# Check for saved settings and prompt to reuse
$SavedSettings = Get-SavedSettings
if ($SavedSettings) {
    Write-Host "`nPrevious settings found:"
    Write-Host "  Output Path : $($SavedSettings.OutputPath)"
    Write-Host "  Hemisphere  : $($SavedSettings.Hemisphere)"
    $UsePrevious = Read-Host "`nUse previous settings? (Y/N)"
    if ($UsePrevious.ToUpper() -eq "Y") {
        $OutputPath = $SavedSettings.OutputPath
        $Hemisphere = $SavedSettings.Hemisphere
        Write-Host "Using previous settings.`n"
    } else {
        Write-Host "Starting fresh.`n"
    }
}

# Get source path
if ($SourcePath -eq $null) {
    if ((Read-Host "Use downloads folder? Y/N") -eq "Y") {
        $SourcePath = "$env:USERPROFILE\downloads"
    } else {
        Write-Host "Prompting for source path..."
        $SourcePath = PickFolder -Description "Select SOURCE folder (photos to rename)"
    }
}
Test-Writable $SourcePath -ErrorAction Stop

# Get output path
if ($OutputPath -eq $null) {
    Write-Host "Prompting for output path..."
    $OutputPath = PickFolder -Description "Select OUTPUT folder"
}
Test-Writable $OutputPath -ErrorAction Stop

# Gather camera, film stock, and hemisphere info
if ($Camera    -eq $null) { $Camera    = (Read-Host "Please provide the camera name").Replace(' ','') }
if ($FilmStock -eq $null) { $FilmStock = (Read-Host "Please provide the film stock name").Replace(' ','') }
if ($Hemisphere -eq $null) {
    do {
        $Input = Read-Host "Are you in the Northern or Southern hemisphere? (N/S)"
        $Hemisphere = switch ($Input.ToUpper()) {
            "N"        { "Northern" }
            "S"        { "Southern" }
            "Northern" { "Northern" }
            "Southern" { "Southern" }
        }
    } while ($Hemisphere -notin @("Northern", "Southern"))
}

# Save settings to registry
Save-Settings -OutputPath $OutputPath -Hemisphere $Hemisphere

# Create working directory if it doesn't already exist
if (-not (Test-Path $WorkingDir)) {
    New-Item $WorkingDir -ItemType Directory -Force | Out-Null
}
if (-not (Test-Path $WorkingDir)) {
    throw "Failed to create working directory: $WorkingDir"
}

# Copy source image files to working directory
Write-Host "Copying files to working directory..."
Get-ChildItem -Path $SourcePath -Recurse | Where-Object {
    -not $_.PSIsContainer -and (Test-IsImage $_.FullName)
} | ForEach-Object {
    Copy-Item $_.FullName -Destination $WorkingDir -ErrorAction Stop
}

# Determine roll number by checking output folder for existing matches
$RollPattern = "*-$($Camera)-$($FilmStock)-R*"
$ExistingRolls = Get-ChildItem -Path $OutputPath -Filter $RollPattern -ErrorAction SilentlyContinue |
    ForEach-Object {
        if ($_.BaseName -match '-R(\d+)-') { [int]$Matches[1] }
    } | Sort-Object -Descending

$RollNumber = if ($ExistingRolls) { $ExistingRolls[0] + 1 } else { 1 }
$RollString = "R$('{0:D2}' -f $RollNumber)"

Write-Host "Roll detected as: $RollString"

# Rename files and copy to output
Write-Host "Renaming and copying files to output..."
$PhotoNumber = 1
Get-ChildItem -Path $WorkingDir | Where-Object {
    -not $_.PSIsContainer -and (Test-IsImage $_.FullName)
} | Sort-Object CreationTime | ForEach-Object {

    # Read creation date and determine season
    $CreationDate = $_.CreationTime
    $Year         = $CreationDate.Year
    $Season       = Get-Season -Date $CreationDate

    # Build naming schema now that all variables are populated
    $Extension   = $_.Extension
    $NewName     = "$($Year)-$($Season)-$($Camera)-$($FilmStock)-$($RollString)-$('{0:D4}' -f $PhotoNumber)$Extension"
    $Destination = Join-Path $OutputPath $NewName

    # Handle duplicate filenames (shouldn't normally happen with roll numbers)
    $Counter = 1
    while (Test-Path $Destination) {
        $NewName     = "$($Year)-$($Season)-$($Camera)-$($FilmStock)-$($RollString)-$('{0:D4}' -f $PhotoNumber)_$Counter$Extension"
        $Destination = Join-Path $OutputPath $NewName
        $Counter++
    }

    Write-Host "  $($_.Name)  ->  $NewName"
    Copy-Item $_.FullName -Destination $Destination -ErrorAction Stop

    $PhotoNumber++
}

# Cleanup working directory
Write-Host "Performing cleanup..."
Remove-Item $WorkingDir -Recurse -Force

Write-Host "`nDone! $($PhotoNumber - 1) file(s) processed -> $OutputPath"
