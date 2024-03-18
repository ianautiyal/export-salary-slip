# Created by: Ajay Nautiyal, <aju@nautiyal.dev>, https://aju.nautiyal.dev
# SPlit and Convert Doc to PDF

Add-Type -AssemblyName System.Windows.Forms
Add-type -AssemblyName Microsoft.Office.Interop.Word

function ConvertToTitleCase {
    param(
        [string]$text
    )

    # Convert text to lowercase
    $lowercase = $text.ToLower()

    # Split the text into an array of words
    $words = $lowercase -split '\s+'

    # Iterate through each word and capitalize the first letter
    $titleCaseWords = foreach ($word in $words) {
        $word.Substring(0, 1).ToUpper() + $word.Substring(1)
    }

    # Join the words back together
    $titleCaseText = $titleCaseWords -join ' '

    return $titleCaseText
}

function ExportSalarySlip {
    param (
        [string]$wdSourceFile,
        [string]$wdExportPath
    )

    $wdExportFormat = 17
    $wdOpenAfterExport = $false
    $wdExportOptimizeFor = 1
    $wdExportRange = 3
    $wdExportItem = 0
    $wdIncludeDocProps = $true
    $wdKeepIRM = $true
    $wdCreateBookmarks = 1
    $wdDocStructureTags = $true
    $wdBitmapMissingFonts = $true
    $wdUseISO19005_1 = $false
    $wdDoNotSaveChanges = 0
    $wdActiveEndPageNumber = 3

    $wdApplication = $null
    $wdDocument = $null

    $employeeCodePattern = 'Emp\. Code:\s*([A-Z0-9-]+)'
    $employeeDepartmentPattern = 'Deptt\. :\s*(.*)'

    $exportedSlipCount = 0

    try {
        $wdApplication = New-Object -ComObject "Word.Application"
        $wdApplication.Visible = $false

        $wdDocument = $wdApplication.Documents.Open($wdSourceFile)

        foreach ($table in $wdDocument.Tables) {
            $employeeCode = $null
            $employeeDepartment = $null

            Write-Host "Exporting page ${page}..."

            $rowCount = $table.Rows.Count
            $columnCount = $table.Columns.Count

            for ($r = 1; $r -le $rowCount; $r++) {
                for ($c = 1; $c -le $columnCount; $c++) {
                    # surround each value with quotes to prevent fields that contain the delimiter character would ruin the csv,
                    # double any double-quotes the value may contain,
                    # remove the control characters (0x0D 0x07) Word appends to the cell text
                    # trim the resulting value from leading or trailing whitespace characters
                    try {
                        $content = ($table.Cell($r, $c).Range.Text -replace '"', '""' -replace '[\x00-\x1F\x7F]').Trim()

                        if ($content -match $employeeCodePattern) {
                            $employeeCode = $Matches[1]
                        }

                        if ($content -match $employeeDepartmentPattern) {
                            $employeeDepartment = $Matches[1]
                        }
                    }
                    catch { }
                }
            }

            if ($null -eq $employeeCode -or $null -eq $employeeDepartment) {
                continue
            }

            $employeeDepartment = ConvertToTitleCase ($employeeDepartment -replace "[^a-zA-Z\s]" -replace '\s+', ' ')

            $page = $Table.Range.Information($wdActiveEndPageNumber);

            Write-Host "Exporting ${employeeCode} - ${employeeDepartment}..."

            # make directory if it doesn't exist
            if (!(Test-Path "${wdExportPath}\${employeeDepartment}")) {
                New-Item -ItemType Directory -Force -Path "${wdExportPath}\${employeeDepartment}" > $null
            }

            $wdExportFile = "${wdExportPath}\${employeeDepartment}\${employeeCode}.pdf"
            $wdStartPage = $page
            $wdEndPage = $page

            $wdDocument.ExportAsFixedFormat(
                $wdExportFile,
                $wdExportFormat,
                $wdOpenAfterExport,
                $wdExportOptimizeFor,
                $wdExportRange,
                $wdStartPage,
                $wdEndPage,
                $wdExportItem,
                $wdIncludeDocProps,
                $wdKeepIRM,
                $wdCreateBookmarks,
                $wdDocStructureTags,
                $wdBitmapMissingFonts,
                $wdUseISO19005_1
            )

            $exportedSlipCount++
        }

        $wshShell = New-Object -ComObject WScript.Shell
        $wshShell.Popup("Done. ${exportedSlipCount} salary slips exported to ${wdExportPath}", 0, "Success", 0)
    }
    catch {
        $wshShell = New-Object -ComObject WScript.Shell
        Write-Host $_.Exception
        $wshShell.Popup($_.Exception.ToString(), 0, "Error", 0)
        $wshShell.Popup("Oops, an error occurred. Please try again. Or contact the developer. <Ajay Nautiyal>", 0, "System Error", 0)
        $wshShell = $null
    }
    finally {
        if ($wdDocument) {
            $wdDocument.Close([ref]$wdDoNotSaveChanges)
            $wdDocument = $null
        }

        if ($wdApplication) {
            $wdApplication.Quit()
            $wdApplication = $null
        }

        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
    }
}

function SelectInputFile {
    $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    $openFileDialog.Title = "Select Master Slary Slip file"
    $openFileDialog.InitialDirectory = [Environment]::GetFolderPath("Desktop")
    $openFileDialog.Filter = "Word Documents (*.docx)|*.docx"
    $openFileDialog.FilterIndex = 1
    $openFileDialog.Multiselect = $false
    
    # Show the File Open Dialog
    $result = $openFileDialog.ShowDialog(
        (New-Object System.Windows.Forms.Form -Property @{TopMost = $true })
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $openFileDialog.FileName
    }

    return $null
}

function SelectOutputDir {
    $folderBrowserDialog = New-Object System.Windows.Forms.FolderBrowserDialog
    $folderBrowserDialog.Description = "Select a destination folder for exported salary slips"

    # Show the Folder Browser Dialog
    $result = $folderBrowserDialog.ShowDialog(
        (New-Object System.Windows.Forms.Form -Property @{TopMost = $true })
    )

    if ($result -eq [System.Windows.Forms.DialogResult]::OK) {
        return $folderBrowserDialog.SelectedPath
    }

    return $null
}

$InputFile = SelectInputFile

if ($null -eq $InputFile) {
    Write-Host "No input file selected. Exiting..."
    exit
}

$OutputDir = SelectOutputDir

if ($null -eq $OutputDir) {
    Write-Host "No output directory selected. Exiting..."
    exit
}

Write-Host "Processing ${InputFile} to ${OutputDir}..."
ExportSalarySlip $InputFile $OutputDir
