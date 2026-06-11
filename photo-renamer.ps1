<#
Trent's Photography Script
Partially written by Claude

This script was written to help with sorting my images by year, season and down to camera and film stock.
I will then sort the folder by creation date.

The intention is to create a yearly photobook, sorted by season. I can also add the camera model and film stock as a note in the book if I want.

Processing Offset:
  Film labs typically take days or weeks to return scans. Without an offset, a roll shot in late December
  but scanned in January would be sorted into the wrong season/year. The offset (in days) is subtracted
  from each file's creation time before the season and year are determined, approximating the original
  shoot date. The offset is saved to the registry alongside other persistent settings.
#>

$SourcePath       = $null
$OutputPath       = $null
$Camera           = $null
$FilmStock        = $null
$Hemisphere       = $null  # Set to "Northern" or "Southern", or leave $null to be prompted
$ProcessingOffset = $null  # Days to subtract from file creation time; leave $null to be prompted

# Other variables you probably don't need to edit
$WorkingDir = "$env:TEMP\photoscript"
$RegPath    = "HKCU:\Software\TrentsPhotoScript"

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

function Save-Settings ([string]$OutputPath, [string]$Hemisphere, [int]$ProcessingOffset) {
    if (-not (Test-Path $RegPath)) {
        New-Item -Path $RegPath -Force | Out-Null
    }
    Set-ItemProperty -Path $RegPath -Name "OutputPath"       -Value $OutputPath
    Set-ItemProperty -Path $RegPath -Name "Hemisphere"       -Value $Hemisphere
    Set-ItemProperty -Path $RegPath -Name "ProcessingOffset" -Value $ProcessingOffset
}

function Prompt-WithDefault ([string]$Prompt, [string]$Default) {
    # Shows the prompt with the saved default in brackets. Pressing Enter accepts the default.
    $Display = if ($Default -ne $null -and $Default -ne "") { "$Prompt [$Default]: " } else { "${Prompt}: " }
    $Input = Read-Host $Display
    if ([string]::IsNullOrWhiteSpace($Input)) { return $Default } else { return $Input }
}

#### ---- START SCRIPT ---- ####

# Check for saved settings and prompt to reuse
$SavedSettings = Get-SavedSettings
if ($SavedSettings) {
    Write-Host "`nPrevious settings found:"
    Write-Host "  Output Path       : $($SavedSettings.OutputPath)"
    Write-Host "  Hemisphere        : $($SavedSettings.Hemisphere)"
    Write-Host "  Processing Offset : $($SavedSettings.ProcessingOffset) day(s)"
    $UsePrevious = Read-Host "`nUse previous settings? (Y/N)"
    if ($UsePrevious.ToUpper() -eq "Y") {
        $OutputPath       = $SavedSettings.OutputPath
        $Hemisphere       = $SavedSettings.Hemisphere
        $ProcessingOffset = [int]$SavedSettings.ProcessingOffset
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

# Gather camera and film stock
if ($Camera    -eq $null) { $Camera    = (Read-Host "Please provide the camera name").Replace(' ','') }
if ($FilmStock -eq $null) { $FilmStock = (Read-Host "Please provide the film stock name").Replace(' ','') }

# Hemisphere — prompt with saved default if available
if ($Hemisphere -eq $null) {
    $HemisphereDefault = if ($SavedSettings.Hemisphere) { $SavedSettings.Hemisphere } else { $null }
    do {
        $HemisphereInput = Prompt-WithDefault -Prompt "Northern or Southern hemisphere? (N/S)" -Default $HemisphereDefault
        $Hemisphere = switch ($HemisphereInput.ToUpper()) {
            "N"        { "Northern" }
            "S"        { "Southern" }
            "Northern" { "Northern" }
            "Southern" { "Southern" }
        }
    } while ($Hemisphere -notin @("Northern", "Southern"))
}

# Processing offset — prompt with saved default, Enter accepts it
if ($ProcessingOffset -eq $null) {
    $OffsetDefault = if ($SavedSettings.ProcessingOffset -ne $null) { "$($SavedSettings.ProcessingOffset)" } else { "30" }
    do {
        $OffsetInput = Prompt-WithDefault -Prompt "Lab processing offset in days (subtracted from scan date to estimate shoot date)" -Default $OffsetDefault
        $ParsedOffset = $null
        $Valid = [int]::TryParse($OffsetInput, [ref]$ParsedOffset) -and $ParsedOffset -ge 0
        if (-not $Valid) { Write-Host "  Please enter a whole number of 0 or more." }
    } while (-not $Valid)
    $ProcessingOffset = $ParsedOffset
}

Write-Host "`nProcessing offset set to $ProcessingOffset day(s). Season and year will be calculated from scan date minus $ProcessingOffset day(s).`n"

# Save settings to registry
Save-Settings -OutputPath $OutputPath -Hemisphere $Hemisphere -ProcessingOffset $ProcessingOffset

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

    # Subtract processing offset to approximate the original shoot date
    $ScanDate  = $_.CreationTime
    $ShootDate = $ScanDate.AddDays(-$ProcessingOffset)
    $Year      = $ShootDate.Year
    $Season    = Get-Season -Date $ShootDate

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
