<#
.SYNOPSIS
    Creates month folders named "1. January", "2. February", etc. 
    in the same directory as this script,
    and creates subfolders (yyyy-MM-dd) for each date matching the chosen Year & DayOfWeek.
#>

try {
    # Load .NET assemblies for GUI
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Host "Error: Unable to load required .NET assemblies (System.Windows.Forms / System.Drawing)."
    Write-Host "Details: $($_.Exception.Message)"
    return
}

function Create-DayFolders {
<#
.SYNOPSIS
    Creates numbered month folders ("1. January", "2. February", etc.)
    in $PSScriptRoot (the script's directory),
    and subfolders named yyyy-MM-dd for each date in the chosen Year
    that falls on the chosen DayOfWeek.
#>
    param(
        [int]$Year,
        [string]$DayOfWeek
    )

    # Parent directory is the script's directory
    $ParentDir = $PSScriptRoot

    # Convert the chosen day string (e.g., "Wednesday") to the actual enum value [System.DayOfWeek]::Wednesday
    try {
        $chosenDayEnum = [System.Enum]::Parse([System.DayOfWeek], $DayOfWeek)
    }
    catch {
        # If parse fails (invalid day name), default to Wednesday
        $chosenDayEnum = [System.DayOfWeek]::Wednesday
    }

    # If the parent directory doesn't exist (unlikely), create it
    if (!(Test-Path $ParentDir)) {
        New-Item -ItemType Directory -Path $ParentDir | Out-Null
    }

    # Create 12 month folders with numbering (1. January, 2. February, etc.)
    # We'll keep the array of month names to help with subfolders for each date
    $months = @(
        'January', 'February', 'March', 'April', 'May', 'June',
        'July', 'August', 'September', 'October', 'November', 'December'
    )

    # For each month 1..12, build a name: "X. MonthName"
    for ($monthNum = 1; $monthNum -le 12; $monthNum++) {
        $monthName      = $months[$monthNum - 1]
        $monthFolderName = "$monthNum. $monthName"
        $monthPath       = Join-Path $ParentDir $monthFolderName

        if (!(Test-Path $monthPath)) {
            New-Item -ItemType Directory -Path $monthPath | Out-Null
        }
    }

    # For each day of each month, check if it is the chosen DayOfWeek
    for ($monthNum = 1; $monthNum -le 12; $monthNum++) {
        $monthName       = $months[$monthNum - 1]
        $monthFolderName = "$monthNum. $monthName"
        $monthPath       = Join-Path $ParentDir $monthFolderName

        $daysInMonth = [DateTime]::DaysInMonth($Year, $monthNum)

        for ($d = 1; $d -le $daysInMonth; $d++) {
            # Use a universal constructor
            $thisDate = [datetime]::ParseExact("$Year-$monthNum-$d","yyyy-M-d",$null)

            if ($thisDate.DayOfWeek -eq $chosenDayEnum) {
                $dateFolder = Join-Path $monthPath $thisDate.ToString('yyyy-MM-dd')
                if (!(Test-Path $dateFolder)) {
                    New-Item -ItemType Directory -Path $dateFolder | Out-Null
                }
            }
        }
    }

    # After creation, show a message box
    [System.Windows.Forms.MessageBox]::Show(
        "All $DayOfWeek folders for $Year have been created in `"$ParentDir`".",
        "Done",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

#
# Build the GUI
#

# Create a form
$form                = New-Object System.Windows.Forms.Form
$form.Text           = "Create Day Folders by Year"
$form.StartPosition  = "CenterScreen"
$form.Size           = New-Object System.Drawing.Size(430, 180)
$form.Topmost        = $true

#
# Label & NumericUpDown for Year
#
$lblYear             = New-Object System.Windows.Forms.Label
$lblYear.Text        = "Year:"
$lblYear.AutoSize    = $true
$lblYear.Location    = New-Object System.Drawing.Point(10, 20)
$form.Controls.Add($lblYear)

$nudYear             = New-Object System.Windows.Forms.NumericUpDown
$nudYear.Location    = New-Object System.Drawing.Point(80, 16)
$nudYear.Size        = New-Object System.Drawing.Size(80, 20)
$nudYear.Minimum     = 1
$nudYear.Maximum     = 9999
$nudYear.Value       = (Get-Date).Year  # default to current year
$form.Controls.Add($nudYear)

#
# Label & ComboBox for Day of the Week
#
$lblDay              = New-Object System.Windows.Forms.Label
$lblDay.Text         = "Day of Week:"
$lblDay.AutoSize     = $true
$lblDay.Location     = New-Object System.Drawing.Point(10, 60)
$form.Controls.Add($lblDay)

$cbDay               = New-Object System.Windows.Forms.ComboBox
$cbDay.DropDownStyle = "DropDownList"
$cbDay.Location      = New-Object System.Drawing.Point(80, 56)
$cbDay.Size          = New-Object System.Drawing.Size(120, 20)
# Valid day names (English) that map to [System.DayOfWeek] Enum
$cbDay.Items.AddRange(@("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday"))
$cbDay.SelectedItem  = "Wednesday"
$form.Controls.Add($cbDay)

#
# Create Folders & Cancel Buttons
#
$btnCreate           = New-Object System.Windows.Forms.Button
$btnCreate.Text      = "Create Folders"
$btnCreate.Location  = New-Object System.Drawing.Point(80, 100)
$btnCreate.Add_Click({
    $chosenYear = [int]$nudYear.Value
    $chosenDay  = $cbDay.SelectedItem

    Create-DayFolders -Year $chosenYear -DayOfWeek $chosenDay
    $form.Close()
})
$form.Controls.Add($btnCreate)

$btnCancel           = New-Object System.Windows.Forms.Button
$btnCancel.Text      = "Cancel"
$btnCancel.Location  = New-Object System.Drawing.Point(200, 100)
$btnCancel.Add_Click({
    $form.Close()
})
$form.Controls.Add($btnCancel)

#
# Show the form
#
$form.ShowDialog() | Out-Null
