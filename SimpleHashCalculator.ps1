<#
.SYNOPSIS
    Calculate file hashes via Windows 'Send to' context menu
.DESCRIPTION
    Computes hashes for selected files and displays results
.NOTES
    Version: 1.4
    Author: Your Name
    Created: $(Get-Date -Format "yyyy-MM-dd")
#>

param(
    [Parameter(Mandatory=$true, ValueFromPipeline=$true, ValueFromRemainingArguments=$true)]
    [Alias('Path')]
    [string[]]$Files
)

# Handle files passed via 'Send to' context menu
if ($Files -eq $null -or $Files.Count -eq 0) {
    [System.Windows.Forms.MessageBox]::Show("Please select files via right-click 'Send to' menu", "Error", "OK", "Error") | Out-Null
    exit
}

# Normalize file paths input to array
$Files = @($Files | Where-Object { $_ -ne $null } | ForEach-Object {
    if ($_ -is [string]) { $_ }
    elseif ($_ -is [array]) { @($_) }
    else { $_.ToString() }
}) | ForEach-Object { $_.Trim('"') }

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Create main form
$form = New-Object System.Windows.Forms.Form
$form.Text = "File Hash Calculator"
$form.Width = 650
$form.Height = 450
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.TopMost = $true

# Hash storage for copying
$hashStorage = @()

# Create hash algorithm selection
$labelAlgorithm = New-Object System.Windows.Forms.Label
$labelAlgorithm.Text = "Hash Algorithm:"
$labelAlgorithm.Location = New-Object System.Drawing.Point(20, 20)
$labelAlgorithm.Width = 100

$comboBoxAlgorithm = New-Object System.Windows.Forms.ComboBox
$comboBoxAlgorithm.Location = New-Object System.Drawing.Point(120, 20)
$comboBoxAlgorithm.Width = 120
$comboBoxAlgorithm.Items.AddRange(@("MD5", "SHA1", "SHA256", "SHA384", "SHA512"))
$comboBoxAlgorithm.SelectedIndex = 2 # Default to SHA256

# Create results textbox
$textBoxResults = New-Object System.Windows.Forms.TextBox
$textBoxResults.Location = New-Object System.Drawing.Point(20, 60)
$textBoxResults.Width = 600
$textBoxResults.Height = 320
$textBoxResults.Multiline = $true
$textBoxResults.ScrollBars = "Vertical"
$textBoxResults.ReadOnly = $true
$textBoxResults.Font = New-Object System.Drawing.Font("Consolas", 10)

# Create buttons
$buttonCopy = New-Object System.Windows.Forms.Button
$buttonCopy.Location = New-Object System.Drawing.Point(500, 20)
$buttonCopy.Text = "Copy All Hashes"
$buttonCopy.Width = 100
$buttonCopy.Add_Click({
    if ($hashStorage.Count -gt 0) {
        try {
            $hashesToCopy = $hashStorage -join "`r`n"
            [System.Windows.Forms.Clipboard]::SetText($hashesToCopy)
            $textBoxResults.AppendText("`r`n`r`nCopied successfully!")
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show("Failed to copy hashes: $_", "Error", "OK", "Error") | Out-Null
        }
    }
})

# Use PowerShell built-in Get-FileHash cmdlet
function Get-FileHashWithProgress {
    param(
        [string]$FilePath,
        [string]$Algorithm
    )
    
    try {
        $hash = Get-FileHash -Path $FilePath -Algorithm $Algorithm -ErrorAction SilentlyContinue        
        if (-not $hash) {
            throw "Failed to calculate hash for $FilePath"
        }
        
        $form.Text = "File Hash Calculator - Processing: 100%"
        [System.Windows.Forms.Application]::DoEvents()
        return $hash.Hash.ToLower()
    }
    catch {
        throw $_
    }
}

# Calculate hashes for all files
function Calculate-Hashes {
    $global:hashStorage = @()
    $currentAlgorithm = $comboBoxAlgorithm.SelectedItem.ToString()
    $fileCount = $Files.Count

    $form.Text = "File Hash Calculator"
    $results = @()
    
    $currentFile = 1
    foreach ($file in $Files) {
        try {
            if (Test-Path $file -PathType Container) {
                $results += "[$currentFile/$fileCount] Skipped folder: $file"
                $currentFile++
                continue
            }

            # Clear memory before processing each file to prevent memory leaks
            [System.GC]::Collect()
            [System.GC]::WaitForPendingFinalizers()

            $hash = Get-FileHashWithProgress -FilePath $file -Algorithm $currentAlgorithm
            $global:hashStorage += $hash  # Store hash for copying

            $results += "[$currentFile/$fileCount] $file"
            $results += "$($currentAlgorithm): $hash"
            $results += "-" * 70
            $currentFile++
        }
        catch {
            $results += "[$currentFile/$fileCount] Error processing $file : $_"
            $currentFile++
            # Continue processing next file even if one fails
            continue
        }
    }

    $textBoxResults.Text = $results -join "`r`n"
    $form.Text = "File Hash Calculator"
}

# Event handlers
$comboBoxAlgorithm.Add_SelectedIndexChanged({
    Calculate-Hashes
})

# Add controls to form
$form.Controls.Add($labelAlgorithm)
$form.Controls.Add($comboBoxAlgorithm)
$form.Controls.Add($textBoxResults)
$form.Controls.Add($buttonCopy)


# Initial calculation
Calculate-Hashes

# Show form
$form.ShowDialog()
