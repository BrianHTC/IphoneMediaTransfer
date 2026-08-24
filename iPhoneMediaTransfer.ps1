# iPhoneMediaTransfer.ps1
# Correct order: 1) source folder -> [n]   2) then Skip/Overwrite/Rename_1

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$form = New-Object System.Windows.Forms.Form
$form.Text = "iPhone Media Transfer"
$form.Size = New-Object System.Drawing.Size(980, 1020)
$form.StartPosition = "CenterScreen"
$form.FormBorderStyle = "FixedDialog"
$form.MaximizeBox = $false
$form.MinimizeBox = $true

# ========== Fonts ==========
$defaultFont = New-Object System.Drawing.Font("Segoe UI", 14)
$logFont     = New-Object System.Drawing.Font("Consolas", 14.5)
$buttonFont  = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)

# ========== Destination ==========
$lblDest = New-Object System.Windows.Forms.Label
$lblDest.Text = "Destination folder:"
$lblDest.Location = New-Object System.Drawing.Point(25, 18)
$lblDest.AutoSize = $true
$lblDest.Font = $defaultFont
$form.Controls.Add($lblDest)

$txtDest = New-Object System.Windows.Forms.TextBox
$txtDest.Location = New-Object System.Drawing.Point(25, 50)
$txtDest.Size = New-Object System.Drawing.Size(720, 32)
$txtDest.Font = $defaultFont
$txtDest.Text = "C:\Users\user\Desktop\iPhoneImport"
$form.Controls.Add($txtDest)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = "Browse..."
$btnBrowse.Location = New-Object System.Drawing.Point(760, 48)
$btnBrowse.Size = New-Object System.Drawing.Size(140, 36)
$btnBrowse.Font = $defaultFont
$form.Controls.Add($btnBrowse)

$lblSpace = New-Object System.Windows.Forms.Label
$lblSpace.Text = "Disk space: (select destination)"
$lblSpace.Location = New-Object System.Drawing.Point(25, 95)
$lblSpace.Size = New-Object System.Drawing.Size(880, 28)
$lblSpace.Font = $defaultFont
$lblSpace.ForeColor = [System.Drawing.Color]::DarkBlue
$form.Controls.Add($lblSpace)

# ========== Sort + Conflict ==========
$lblSort = New-Object System.Windows.Forms.Label
$lblSort.Text = "Sort folders by:"
$lblSort.Location = New-Object System.Drawing.Point(25, 140)
$lblSort.AutoSize = $true
$lblSort.Font = $defaultFont
$form.Controls.Add($lblSort)

$cmbSort = New-Object System.Windows.Forms.ComboBox
$cmbSort.Location = New-Object System.Drawing.Point(25, 175)
$cmbSort.Size = New-Object System.Drawing.Size(300, 32)
$cmbSort.DropDownStyle = "DropDownList"
$cmbSort.Font = $defaultFont
$cmbSort.Items.AddRange(@(
    "Year only (YYYY)",
    "Year\Month (YYYY\MM)",
    "Year\Month\Day (YYYY\MM\DD)"
))
$cmbSort.SelectedIndex = 1
$form.Controls.Add($cmbSort)

$lblConflict = New-Object System.Windows.Forms.Label
$lblConflict.Text = "File conflict:"
$lblConflict.Location = New-Object System.Drawing.Point(350, 140)
$lblConflict.AutoSize = $true
$lblConflict.Font = $defaultFont
$form.Controls.Add($lblConflict)

$cmbConflict = New-Object System.Windows.Forms.ComboBox
$cmbConflict.Location = New-Object System.Drawing.Point(350, 175)
$cmbConflict.Size = New-Object System.Drawing.Size(300, 32)
$cmbConflict.DropDownStyle = "DropDownList"
$cmbConflict.Font = $defaultFont
$cmbConflict.Items.AddRange(@("Skip", "Overwrite", "Rename with _1, _2..."))
$cmbConflict.SelectedIndex = 0
$form.Controls.Add($cmbConflict)

# ========== Date range filter ==========
$grpRange = New-Object System.Windows.Forms.GroupBox
$grpRange.Text = "Import date range filter (optional)"
$grpRange.Location = New-Object System.Drawing.Point(25, 255)
$grpRange.Size = New-Object System.Drawing.Size(900, 120)
$grpRange.Font = $defaultFont
$form.Controls.Add($grpRange)

$chkEnableRange = New-Object System.Windows.Forms.CheckBox
$chkEnableRange.Text = "Enable date range filter"
$chkEnableRange.Location = New-Object System.Drawing.Point(20, 30)
$chkEnableRange.AutoSize = $true
$chkEnableRange.Font = $defaultFont
$grpRange.Controls.Add($chkEnableRange)

$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = "From:"
$lblFrom.Location = New-Object System.Drawing.Point(20, 70)
$lblFrom.AutoSize = $true
$lblFrom.Font = $defaultFont
$grpRange.Controls.Add($lblFrom)

$dtFrom = New-Object System.Windows.Forms.DateTimePicker
$dtFrom.Location = New-Object System.Drawing.Point(90, 65)
$dtFrom.Size = New-Object System.Drawing.Size(180, 32)
$dtFrom.Format = "Custom"
$dtFrom.CustomFormat = "yyyy/MM/dd"
$dtFrom.Font = $defaultFont
$dtFrom.Enabled = $false
$grpRange.Controls.Add($dtFrom)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = "To:"
$lblTo.Location = New-Object System.Drawing.Point(290, 70)
$lblTo.AutoSize = $true
$lblTo.Font = $defaultFont
$grpRange.Controls.Add($lblTo)

$dtTo = New-Object System.Windows.Forms.DateTimePicker
$dtTo.Location = New-Object System.Drawing.Point(340, 65)
$dtTo.Size = New-Object System.Drawing.Size(180, 32)
$dtTo.Format = "Custom"
$dtTo.CustomFormat = "yyyy/MM/dd"
$dtTo.Font = $defaultFont
$dtTo.Enabled = $false
$grpRange.Controls.Add($dtTo)

$lblRangeHint = New-Object System.Windows.Forms.Label
$lblRangeHint.Text = "Tip: Set From = To for a single day"
$lblRangeHint.Location = New-Object System.Drawing.Point(540, 70)
$lblRangeHint.Size = New-Object System.Drawing.Size(340, 28)
$lblRangeHint.Font = $defaultFont
$lblRangeHint.ForeColor = [System.Drawing.Color]::Gray
$grpRange.Controls.Add($lblRangeHint)

$chkEnableRange.Add_CheckedChanged({
    $dtFrom.Enabled = $chkEnableRange.Checked
    $dtTo.Enabled   = $chkEnableRange.Checked
})

# ========== File type filter ==========
$grpType = New-Object System.Windows.Forms.GroupBox
$grpType.Text = "File type filter"
$grpType.Location = New-Object System.Drawing.Point(25, 395)
$grpType.Size = New-Object System.Drawing.Size(900, 150)
$grpType.Font = $defaultFont
$form.Controls.Add($grpType)

$chkPhotos = New-Object System.Windows.Forms.CheckBox
$chkPhotos.Text = "Photos (jpg, jpeg, png, gif, bmp, tif, tiff)"
$chkPhotos.Location = New-Object System.Drawing.Point(20, 35)
$chkPhotos.AutoSize = $true
$chkPhotos.Font = $defaultFont
$chkPhotos.Checked = $true
$grpType.Controls.Add($chkPhotos)

$chkHeic = New-Object System.Windows.Forms.CheckBox
$chkHeic.Text = "HEIC / HEIF"
$chkHeic.Location = New-Object System.Drawing.Point(480, 35)
$chkHeic.AutoSize = $true
$chkHeic.Font = $defaultFont
$chkHeic.Checked = $true
$grpType.Controls.Add($chkHeic)

$chkVideos = New-Object System.Windows.Forms.CheckBox
$chkVideos.Text = "Videos (mov, mp4, m4v, avi, mkv, 3gp, wmv)"
$chkVideos.Location = New-Object System.Drawing.Point(20, 75)
$chkVideos.AutoSize = $true
$chkVideos.Font = $defaultFont
$chkVideos.Checked = $true
$grpType.Controls.Add($chkVideos)

$chkLive = New-Object System.Windows.Forms.CheckBox
$chkLive.Text = "Live Photo companions (.aae)"
$chkLive.Location = New-Object System.Drawing.Point(480, 75)
$chkLive.AutoSize = $true
$chkLive.Font = $defaultFont
$chkLive.Checked = $false
$grpType.Controls.Add($chkLive)

$lblCustom = New-Object System.Windows.Forms.Label
$lblCustom.Text = "Custom extensions (e.g. .pdf,.doc,.zip):"
$lblCustom.Location = New-Object System.Drawing.Point(20, 115)
$lblCustom.AutoSize = $true
$lblCustom.Font = $defaultFont
$grpType.Controls.Add($lblCustom)

$txtCustomExt = New-Object System.Windows.Forms.TextBox
$txtCustomExt.Location = New-Object System.Drawing.Point(420, 110)
$txtCustomExt.Size = New-Object System.Drawing.Size(450, 32)
$txtCustomExt.Font = $defaultFont
$grpType.Controls.Add($txtCustomExt)

# ========== Other options ==========
$chkRecursive = New-Object System.Windows.Forms.CheckBox
$chkRecursive.Text = "Include all subfolders under Internal Storage / DCIM"
$chkRecursive.Location = New-Object System.Drawing.Point(25, 560)
$chkRecursive.AutoSize = $true
$chkRecursive.Font = $defaultFont
$chkRecursive.Checked = $true
$form.Controls.Add($chkRecursive)

$chkVerify = New-Object System.Windows.Forms.CheckBox
$chkVerify.Text = "Verify after copy (size + duration for video/audio)"
$chkVerify.Location = New-Object System.Drawing.Point(25, 600)
$chkVerify.AutoSize = $true
$chkVerify.Font = $defaultFont
$chkVerify.Checked = $true
$form.Controls.Add($chkVerify)

# ========== Progress + Log ==========
$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(25, 650)
$progress.Size = New-Object System.Drawing.Size(900, 28)
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready. Connect & unlock iPhone, then click Start."
$lblStatus.Location = New-Object System.Drawing.Point(25, 690)
$lblStatus.Size = New-Object System.Drawing.Size(900, 28)
$lblStatus.Font = $defaultFont
$form.Controls.Add($lblStatus)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.Location = New-Object System.Drawing.Point(25, 725)
$txtLog.Size = New-Object System.Drawing.Size(900, 170)
$txtLog.ReadOnly = $true
$txtLog.Font = $logFont
$form.Controls.Add($txtLog)

# ========== Buttons ==========
$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "Start Copy"
$btnStart.Location = New-Object System.Drawing.Point(25, 920)
$btnStart.Size = New-Object System.Drawing.Size(180, 45)
$btnStart.Font = $buttonFont
$btnStart.BackColor = [System.Drawing.Color]::LightGreen
$form.Controls.Add($btnStart)

$btnAbort = New-Object System.Windows.Forms.Button
$btnAbort.Text = "Abort"
$btnAbort.Location = New-Object System.Drawing.Point(230, 920)
$btnAbort.Size = New-Object System.Drawing.Size(180, 45)
$btnAbort.Font = $buttonFont
$btnAbort.BackColor = [System.Drawing.Color]::LightCoral
$btnAbort.Enabled = $false
$form.Controls.Add($btnAbort)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Text = "Close"
$btnClose.Location = New-Object System.Drawing.Point(745, 920)
$btnClose.Size = New-Object System.Drawing.Size(180, 45)
$btnClose.Font = $buttonFont
$btnClose.Add_Click({ $form.Close() })
$form.Controls.Add($btnClose)

# Global abort and conditional staging state
$script:abortRequested = $false
$script:stageRoot = Join-Path $env:LOCALAPPDATA "iPhoneMediaCopy\Staging"
$script:currentStageDir = $null

function Remove-IPhoneStaging {
    if ($script:currentStageDir -and (Test-Path -LiteralPath $script:currentStageDir)) {
        Remove-Item -LiteralPath $script:currentStageDir -Recurse -Force -ErrorAction SilentlyContinue
    }
    $script:currentStageDir = $null

    if (Test-Path -LiteralPath $script:stageRoot) {
        Get-ChildItem -LiteralPath $script:stageRoot -Force -ErrorAction SilentlyContinue |
            Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
        Remove-Item -LiteralPath $script:stageRoot -Recurse -Force -ErrorAction SilentlyContinue
    }

    $stageParent = Split-Path $script:stageRoot -Parent
    if (Test-Path -LiteralPath $stageParent) {
        $remaining = @(Get-ChildItem -LiteralPath $stageParent -Force -ErrorAction SilentlyContinue)
        if ($remaining.Count -eq 0) {
            Remove-Item -LiteralPath $stageParent -Force -ErrorAction SilentlyContinue
        }
    }
}

$btnAbort.Add_Click({
    $script:abortRequested = $true
    Write-Log ">>> ABORT requested - stopping as soon as possible..."
    $lblStatus.Text = "ABORTING NOW..."
    $btnAbort.Enabled = $false
    $btnStart.Enabled = $false
    Remove-IPhoneStaging
})

# ========== Disk space helper ==========
function Update-DiskSpace {
    try {
        $path = $txtDest.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($path)) {
            $lblSpace.Text = "Disk space: (select destination)"
            return
        }
        $drive = [System.IO.Path]::GetPathRoot($path)
        if (-not $drive) {
            $lblSpace.Text = "Disk space: invalid path"
            return
        }
        $driveInfo = New-Object System.IO.DriveInfo($drive)
        if (-not $driveInfo.IsReady) {
            $lblSpace.Text = "Disk space: drive not ready"
            return
        }
        $freeGB  = [math]::Round($driveInfo.AvailableFreeSpace / 1GB, 2)
        $totalGB = [math]::Round($driveInfo.TotalSize / 1GB, 2)
        $usedGB  = [math]::Round(($driveInfo.TotalSize - $driveInfo.AvailableFreeSpace) / 1GB, 2)
        $pct     = [math]::Round(($driveInfo.AvailableFreeSpace / $driveInfo.TotalSize) * 100, 1)

        $lblSpace.Text = "Disk space on $drive  ->  Free: $freeGB GB   |   Used: $usedGB GB   |   Total: $totalGB GB   ($pct% free)"
        
        if ($freeGB -lt 5) {
            $lblSpace.ForeColor = [System.Drawing.Color]::Red
        } elseif ($freeGB -lt 20) {
            $lblSpace.ForeColor = [System.Drawing.Color]::DarkOrange
        } else {
            $lblSpace.ForeColor = [System.Drawing.Color]::DarkGreen
        }
    } catch {
        $lblSpace.Text = "Disk space: cannot read"
        $lblSpace.ForeColor = [System.Drawing.Color]::Red
    }
}

$btnBrowse.Add_Click({
    $fb = New-Object System.Windows.Forms.FolderBrowserDialog
    $fb.Description = "Select destination folder"
    if ($fb.ShowDialog() -eq "OK") {
        $txtDest.Text = $fb.SelectedPath
        Update-DiskSpace
    }
})
$txtDest.Add_TextChanged({ Update-DiskSpace })
Update-DiskSpace

# ========== Helper functions ==========
function Write-Log($msg) {
    $timestamp = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("$timestamp  $msg`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-PhoneRoot {
    try {
        $shell = New-Object -ComObject Shell.Application
        $thisPC = $shell.NameSpace("shell:MyComputerFolder")
        $phone = $thisPC.Items() | Where-Object {
            $_.Name -match "iPhone|Apple" -or $_.Type -match "Portable Device|Mobile|Portable Media"
        } | Select-Object -First 1
        if (-not $phone) { return $null }
        return $phone.GetFolder()
    } catch {
        Write-Log "ERROR Get-PhoneRoot: $_"
        return $null
    }
}

function Get-FolderByPath($parentFolder, $relativePath) {
    try {
        $current = $parentFolder
        foreach ($part in ($relativePath -split '\\' | Where-Object { $_ })) {
            $item = $current.Items() | Where-Object { $_.IsFolder -and $_.Name -eq $part } | Select-Object -First 1
            if (-not $item) { return $null }
            $current = $item.GetFolder()
        }
        return $current
    } catch {
        Write-Log "ERROR Get-FolderByPath: $_"
        return $null
    }
}

function Get-MediaDate($item) {
    try {
        $d = $item.ExtendedProperty("System.DateTaken")
        if ($d) { return [datetime]$d }
    } catch {}
    try {
        $d = $item.ExtendedProperty("System.DateModified")
        if ($d -and ([datetime]$d).Year -gt 1900) { return [datetime]$d }
    } catch {}
    try {
        $d = $item.ModifyDate
        if ($d -and $d.Year -gt 1900) { return $d }
    } catch {}
    return $null
}

function Build-AllowedExtensions {
    $exts = New-Object System.Collections.Generic.HashSet[string]
    if ($chkPhotos.Checked) {
        @('.jpg','.jpeg','.png','.gif','.bmp','.tif','.tiff') | ForEach-Object { [void]$exts.Add($_) }
    }
    if ($chkHeic.Checked) {
        @('.heic','.heif') | ForEach-Object { [void]$exts.Add($_) }
    }
    if ($chkVideos.Checked) {
        @('.mov','.mp4','.m4v','.avi','.mkv','.3gp','.wmv') | ForEach-Object { [void]$exts.Add($_) }
    }
    if ($chkLive.Checked) {
        [void]$exts.Add('.aae')
    }
    $custom = $txtCustomExt.Text.Trim()
    if ($custom) {
        $custom -split ',' | ForEach-Object {
            $e = $_.Trim().ToLower()
            if ($e -and -not $e.StartsWith('.')) { $e = ".$e" }
            if ($e) { [void]$exts.Add($e) }
        }
    }
    return $exts
}

function Get-SourceFolderSuffixNumber($folderName) {
    try {
        $name = ([string]$folderName).Trim()
        if ($name -match '(?i)_([b-z])$') {
            $letter = $Matches[1].ToLowerInvariant()
            return ([int][char]$letter) - ([int][char]'a')
        }
    } catch {
        Write-Log "WARNING suffix detection for folder '$folderName': $_"
    }
    return 0
}
function Get-UniqueName($dir, $name) {
    $base = [IO.Path]::GetFileNameWithoutExtension($name)
    $ext  = [IO.Path]::GetExtension($name)
    $i = 1
    do {
        $candidate = Join-Path $dir "${base}_${i}${ext}"
        $i++
    } while (Test-Path $candidate)
    return (Split-Path $candidate -Leaf)
}

function Get-ShellSize($shellItem) {
    try {
        $s = $shellItem.ExtendedProperty("System.Size")
        if ($s -ne $null) { return [int64]$s }
    } catch {}
    try {
        if ($shellItem.Size -gt 0) { return [int64]$shellItem.Size }
    } catch {}
    return -1
}

function Get-LocalFileSize($path) {
    try {
        return (Get-Item -LiteralPath $path -ErrorAction Stop).Length
    } catch { return -1 }
}

function Get-MediaDurationSeconds($path) {
    try {
        $shell = New-Object -ComObject Shell.Application
        $folder = $shell.NameSpace((Split-Path $path -Parent))
        $file   = $folder.ParseName((Split-Path $path -Leaf))
        if (-not $file) { return -1 }
        $dur = $file.ExtendedProperty("System.Media.Duration")
        if ($dur -ne $null -and $dur -gt 0) {
            return [math]::Round($dur / 10000000.0, 3)
        }
    } catch {}
    return -1
}

function Is-VideoOrAudio($ext) {
    $media = @('.mov','.mp4','.m4v','.avi','.mkv','.3gp','.wmv','.mp3','.m4a','.aac','.wav','.flac')
    return $media -contains $ext.ToLower()
}

function Test-SourceFolderInDateRange($folderName, $useRange, $fromDate, $toDate) {
    if (-not $useRange) { return $true }

    $name = ([string]$folderName).Trim()

    # iPhone folders such as 202207, 202207_a, 202207_b, and so on.
    if ($name -match '^(\d{4})(\d{2})(?:_[a-z])?$') {
        $year = [int]$Matches[1]
        $month = [int]$Matches[2]

        try {
            $folderMonthStart = Get-Date -Year $year -Month $month -Day 1
            $folderMonthEnd = $folderMonthStart.AddMonths(1).AddDays(-1)
            return ($folderMonthEnd.Date -ge $fromDate.Date -and $folderMonthStart.Date -le $toDate.Date)
        } catch {
            return $true
        }
    }

    # Structural folders such as Internal Storage and DCIM must still be entered.
    return $true
}

# ========== Start button ==========
$btnStart.Add_Click({
    $btnStart.Enabled = $false
    $btnAbort.Enabled = $true
    $script:abortRequested = $false
    Remove-IPhoneStaging
    $progress.Value = 0
    $txtLog.Clear()
    $lblStatus.Text = "Working..."

    try {
        Write-Log "Looking for Apple iPhone..."

        $phoneRoot = Get-PhoneRoot
        if (-not $phoneRoot) {
            $msg = "Apple iPhone not found under This PC.`r`nUnlock the phone and tap Trust."
            Write-Log $msg
            $lblStatus.Text = "ERROR: iPhone not found"
            [System.Windows.Forms.MessageBox]::Show($msg, "Error", "OK", "Error")
            return
        }
        Write-Log "Found device: $($phoneRoot.Title)"

        $internal = Get-FolderByPath $phoneRoot "Internal Storage"
        if (-not $internal) {
            Write-Log "Internal Storage not found - using device root"
            $internal = $phoneRoot
        } else {
            Write-Log "Using path: Internal Storage"
        }

        $destRoot = $txtDest.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($destRoot)) { throw "Destination folder is empty" }
        if (-not (Test-Path $destRoot)) {
            New-Item -ItemType Directory -Path $destRoot -Force | Out-Null
            Write-Log "Created destination: $destRoot"
        }

        Update-DiskSpace

        $sortMode     = $cmbSort.SelectedIndex
        $conflictMode = $cmbConflict.SelectedIndex
        $recursive    = $chkRecursive.Checked
        $useRange     = $chkEnableRange.Checked
        $fromDate     = $dtFrom.Value.Date
        $toDate       = $dtTo.Value.Date
        $doVerify     = $chkVerify.Checked

        if ($useRange -and $fromDate -gt $toDate) { throw "From date cannot be later than To date" }

        $allowedExt = Build-AllowedExtensions
        if ($allowedExt.Count -eq 0) { throw "No file types selected." }

        Write-Log "Allowed extensions: $($allowedExt -join ', ')"
        if ($doVerify) { Write-Log "Verification: ENABLED" } else { Write-Log "Verification: disabled" }
        if ($useRange) {
            Write-Log "Date range: $($fromDate.ToString('yyyy/MM/dd')) -> $($toDate.ToString('yyyy/MM/dd'))"
        }

        # Collect files
        $files = New-Object System.Collections.ArrayList
        $foldersToScan = New-Object System.Collections.Queue
        $foldersToScan.Enqueue($internal)

        Write-Log "Scanning folders..."
        if ($useRange) {
            Write-Log "Fast scan: source month folders outside the selected range will be skipped."
        }

        while ($foldersToScan.Count -gt 0) {
            if ($script:abortRequested) { break }
            $folder = $foldersToScan.Dequeue()
            try {
                foreach ($item in @($folder.Items())) {
                    if ($item.IsFolder) {
                        $isKnownMediaFolder = $item.Name -match '^(?i:DCIM|Camera|Photo|Pictures)$'
                        $isIPhoneMonthFolder = $item.Name -match '^\d{6}(?:_[a-z])?$'
                        $folderAllowedByDate = Test-SourceFolderInDateRange $item.Name $useRange $fromDate $toDate

                        if (($recursive -or $isKnownMediaFolder -or $isIPhoneMonthFolder) -and $folderAllowedByDate) {
                            $foldersToScan.Enqueue($item.GetFolder())
                        } elseif ($useRange -and $isIPhoneMonthFolder) {
                            Write-Log "Skip source folder outside date range: $($item.Name)"
                        }
                    } else {
                        $ext = [IO.Path]::GetExtension($item.Name).ToLower()
                        if ($allowedExt.Contains($ext)) {
                            [void]$files.Add([PSCustomObject]@{
                                Item       = $item
                                FolderName = $folder.Title
                            })
                        }
                    }
                }
            } catch { Write-Log "WARNING scanning: $_" }
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($script:abortRequested) {
            Write-Log ">>> Aborted during scan."
            $lblStatus.Text = "Aborted during scan"
            return
        }

        Write-Log "Found $($files.Count) candidate files."

        if ($useRange) {
            $filtered = New-Object System.Collections.ArrayList
            $noDateCount = 0
            foreach ($fileInfo in $files) {
                $d = Get-MediaDate $fileInfo.Item
                if ($null -eq $d) {
                    $noDateCount++
                    continue
                }
                if ($d.Date -ge $fromDate -and $d.Date -le $toDate) {
                    [void]$filtered.Add($fileInfo)
                }
            }
            if ($noDateCount -gt 0) {
                Write-Log "Files without a readable media date: $noDateCount"
            }
            $files = $filtered
            Write-Log "After date filter: $($files.Count) files."
        }

        $total = $files.Count
        if ($total -eq 0) {
            $lblStatus.Text = "No matching files found."
            Write-Log "No files matched."
            return
        }

        $progress.Maximum = $total
        $copied = 0
        $skipped = 0
        $verifiedOk = 0
        $verifyFail = 0
        $errors = 0
        $shell = New-Object -ComObject Shell.Application

        foreach ($fileInfo in $files) {
            $item = $fileInfo.Item
            if ($script:abortRequested) {
                Write-Log ">>> Aborted by user."
                break
            }

            try {
                $date = Get-MediaDate $item
                if ($null -eq $date) { $date = Get-Date }

                switch ($sortMode) {
                    0 { $sub = $date.ToString("yyyy") }
                    1 { $sub = $date.ToString("yyyy\\MM") }
                    2 { $sub = $date.ToString("yyyy\\MM\\dd") }
                }

                $targetDir = Join-Path $destRoot $sub
                if (-not (Test-Path $targetDir)) {
                    New-Item -ItemType Directory -Path $targetDir -Force | Out-Null
                }

                # ============================================================
                # 1. FIRST decide the base name according to source folder
                # ============================================================
                $ext = [IO.Path]::GetExtension($item.Name)
                $baseName = [IO.Path]::GetFileNameWithoutExtension($item.Name)
                $n = Get-SourceFolderSuffixNumber $fileInfo.FolderName   # 0 = normal, 1=_b, 2=_c...

                if ($ext -ieq ".aae") {
                    $proposedName = $item.Name
                }
                elseif ($n -gt 0) {
                    # Source folder ends with _b/_c/... -> always add [n]
                    $proposedName = "${baseName}[$n]${ext}"
                }
                else {
                    $proposedName = $item.Name
                }

                # ============================================================
                # 2. THEN apply conflict rule on the proposed name
                # ============================================================
                $finalName = $proposedName
                $targetFile = Join-Path $targetDir $finalName
                $doCopy = $true

                Write-Log "Source: $($fileInfo.FolderName)\$($item.Name) | Proposed: $proposedName | FolderSuffix=$n"

                $hadConflict = Test-Path -LiteralPath $targetFile

                if ($hadConflict) {
                    switch ($conflictMode) {
                        0 {  # Skip
                            $doCopy = $false
                            $skipped++
                            Write-Log "Skip (exists): source=$($fileInfo.FolderName)\$($item.Name) | destination=$sub\$finalName"
                        }
                        1 {  # Overwrite
                            Write-Log "Overwrite: source=$($fileInfo.FolderName)\$($item.Name) | destination=$sub\$finalName"
                        }
                        2 {  # Rename with _1, _2...
                            $finalName = Get-UniqueName $targetDir $proposedName
                            $targetFile = Join-Path $targetDir $finalName
                            Write-Log "Rename conflict: source=$($fileInfo.FolderName)\$($item.Name) | destination=$sub\$finalName"
                        }
                    }
                }

                if ($doCopy) {
                    $origSize = Get-ShellSize $item
                    $useStaging = ($n -gt 0) -or (($conflictMode -eq 2) -and $hadConflict)

                    if ($useStaging) {
                        if (-not (Test-Path -LiteralPath $script:stageRoot)) {
                            New-Item -ItemType Directory -Path $script:stageRoot -Force | Out-Null
                        }

                        $script:currentStageDir = Join-Path $script:stageRoot ([guid]::NewGuid().ToString('N'))
                        New-Item -ItemType Directory -Path $script:currentStageDir -Force | Out-Null

                        try {
                            $stageNS = $shell.NameSpace($script:currentStageDir)
                            if (-not $stageNS) { throw "Cannot open staging folder: $($script:currentStageDir)" }
                            $stageNS.CopyHere($item, 16)
                            $stageFile = Join-Path $script:currentStageDir $item.Name

                            $waitCount = 0
                            $lastSize = -1
                            $stableCount = 0
                            while ($waitCount -lt 1200) {
                                if ($script:abortRequested) { break }
                                if (Test-Path -LiteralPath $stageFile) {
                                    $currentSize = Get-LocalFileSize $stageFile
                                    if ($currentSize -gt 0 -and $currentSize -eq $lastSize) { $stableCount++ } else { $stableCount = 0 }
                                    $lastSize = $currentSize
                                    if (($origSize -gt 0 -and $currentSize -eq $origSize) -or ($origSize -le 0 -and $stableCount -ge 4)) { break }
                                }
                                Start-Sleep -Milliseconds 500
                                $waitCount++
                                [System.Windows.Forms.Application]::DoEvents()
                            }

                            if ($script:abortRequested) { throw "Copy aborted by user" }
                            if (-not (Test-Path -LiteralPath $stageFile)) { throw "Staged file did not appear before timeout: $($item.Name)" }
                            if ($origSize -gt 0 -and (Get-LocalFileSize $stageFile) -ne $origSize) { throw "Staged file size did not reach source size: $($item.Name)" }

                            if ($conflictMode -eq 1 -and (Test-Path -LiteralPath $targetFile)) {
                                Remove-Item -LiteralPath $targetFile -Force -ErrorAction Stop
                            }
                            Move-Item -LiteralPath $stageFile -Destination $targetFile -Force -ErrorAction Stop
                        }
                        finally {
                            Remove-IPhoneStaging
                        }
                    } else {
                        if ($conflictMode -eq 1 -and $hadConflict) {
                            Remove-Item -LiteralPath $targetFile -Force -ErrorAction Stop
                        }

                        $destNS = $shell.NameSpace($targetDir)
                        if (-not $destNS) { throw "Cannot open destination folder: $targetDir" }
                        $destNS.CopyHere($item, 16)

                        $waitCount = 0
                        $lastSize = -1
                        $stableCount = 0
                        while ($waitCount -lt 1200) {
                            if ($script:abortRequested) { break }
                            if (Test-Path -LiteralPath $targetFile) {
                                $currentSize = Get-LocalFileSize $targetFile
                                if ($currentSize -gt 0 -and $currentSize -eq $lastSize) { $stableCount++ } else { $stableCount = 0 }
                                $lastSize = $currentSize
                                if (($origSize -gt 0 -and $currentSize -eq $origSize) -or ($origSize -le 0 -and $stableCount -ge 4)) { break }
                            }
                            Start-Sleep -Milliseconds 500
                            $waitCount++
                            [System.Windows.Forms.Application]::DoEvents()
                        }

                        if ($script:abortRequested) {
                            if (Test-Path -LiteralPath $targetFile) { Remove-Item -LiteralPath $targetFile -Force -ErrorAction SilentlyContinue }
                            throw "Copy aborted by user"
                        }
                        if (-not (Test-Path -LiteralPath $targetFile)) { throw "Copied file did not appear before timeout: $($item.Name)" }
                        if ($origSize -gt 0 -and (Get-LocalFileSize $targetFile) -ne $origSize) {
                            Remove-Item -LiteralPath $targetFile -Force -ErrorAction SilentlyContinue
                            throw "Copied file size did not reach source size: $($item.Name)"
                        }
                    }

                    if ($script:abortRequested) {
                        Write-Log ">>> Aborted during copy of $($item.Name)"
                        break
                    }

                    $copied++
                    $logMsg = "Copied: $($item.Name) -> $sub\$finalName"
                    if ($n -gt 0) { $logMsg += "  (source folder _$([char]([int][char]'a' + $n)))" }

                    if ($doVerify) {
                        $localSize = Get-LocalFileSize $targetFile
                        $sizeOk = ($origSize -gt 0 -and $localSize -eq $origSize)

                        $durOk = $true
                        $durMsg = ""
                        if (Is-VideoOrAudio $ext) {
                            $localDur = Get-MediaDurationSeconds $targetFile
                            if ($localDur -le 0) {
                                $durOk = $false
                                $durMsg = " | duration unavailable/zero"
                            } else {
                                $durMsg = " | duration $($localDur)s"
                            }
                        }

                        if ($sizeOk -and $durOk) {
                            $verifiedOk++
                            $logMsg += "  [VERIFY OK size=$localSize$durMsg]"
                        } else {
                            $verifyFail++
                            $reason = @()
                            if (-not $sizeOk) { $reason += "size mismatch" }
                            if (-not $durOk)  { $reason += "duration problem" }
                            $logMsg += "  [VERIFY FAIL $($reason -join ', ')]"
                        }
                    }

                    Write-Log $logMsg
                }
            } catch {
                $errors++
                Write-Log "ERROR on '$($item.Name)': $_"
            }

            $progress.Value = [Math]::Min($copied + $skipped + $errors, $total)
            $lblStatus.Text = "Progress: $($progress.Value)/$total   Copied:$copied  Skipped:$skipped  VerifyOK:$verifiedOk  Fail:$verifyFail  Errors:$errors"
            [System.Windows.Forms.Application]::DoEvents()
        }

        if ($script:abortRequested) {
            Write-Log "========== ABORTED =========="
            $lblStatus.Text = "Aborted. Copied $copied files."
        } else {
            Write-Log "========== Finished =========="
            $lblStatus.Text = "Done. Copied $copied files."
        }

        Write-Log "Copied      : $copied"
        Write-Log "Skipped     : $skipped"
        if ($doVerify) {
            Write-Log "Verify OK   : $verifiedOk"
            Write-Log "Verify FAIL : $verifyFail"
        }
        Write-Log "Errors      : $errors"

        Update-DiskSpace

        $summary = "Copied : $copied`nSkipped : $skipped`n"
        if ($doVerify) { $summary += "Verify OK : $verifiedOk`nVerify FAIL : $verifyFail`n" }
        $summary += "Errors : $errors"

        [System.Windows.Forms.MessageBox]::Show($summary, "Complete", "OK", "Information")
    }
    catch {
        Write-Log "FATAL ERROR: $_"
        $lblStatus.Text = "FATAL ERROR - see log"
        [System.Windows.Forms.MessageBox]::Show("Fatal error:`n$_", "Error", "OK", "Error")
    }
    finally {
        Remove-IPhoneStaging
        $btnStart.Enabled = $true
        $btnAbort.Enabled = $false
        $script:abortRequested = $false
        Write-Log "Script finished. Window will close when you click Close."
    }
})

[void]$form.ShowDialog()
exit