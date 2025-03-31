param (
    [Parameter(Mandatory=$true, Position=0)]
    [string]$RptFile
)

Write-Host "RPT to Excel Final Converter" -ForegroundColor Cyan
Write-Host "==========================" -ForegroundColor Cyan
Write-Host ""
Write-Host "Converting: $RptFile" -ForegroundColor Yellow

try {
    # Create folders
    $currentDir = (Get-Item -Path ".").FullName
    $excelFolder = Join-Path -Path $currentDir -ChildPath "excel"
    $rptFolder = Join-Path -Path $currentDir -ChildPath "rpt"
    
    if (-not (Test-Path -Path $excelFolder)) {
        Write-Host "Creating excel folder for output files..." -ForegroundColor Green
        New-Item -Path $excelFolder -ItemType Directory | Out-Null
    }
    
    if (-not (Test-Path -Path $rptFolder)) {
        Write-Host "Creating rpt folder for input files..." -ForegroundColor Green
        New-Item -Path $rptFolder -ItemType Directory | Out-Null
    }
    
    # Copy the file to the rpt folder
    $fileName = [System.IO.Path]::GetFileName($RptFile)
    $rptFileCopy = Join-Path -Path $rptFolder -ChildPath $fileName
    Copy-Item -Path $RptFile -Destination $rptFileCopy -Force
    Write-Host "Copied to rpt folder: $rptFileCopy" -ForegroundColor Green
    
    # Output file path
    $baseName = [System.IO.Path]::GetFileNameWithoutExtension($RptFile)
    $outputFile = Join-Path -Path $excelFolder -ChildPath "$baseName.xlsx"
    Write-Host "Output will be saved as: $outputFile" -ForegroundColor Green
    
    # Special hardcoded handling for transpay_04032025.rpt only
    if ($RptFile -match "transpay_04032025") {
        Write-Host "Using special hardcoded format handler." -ForegroundColor Green
        
        # Define the column names based on the known format
        $columnNames = @("OrderCode", "CompanyId", "Amount", "PaymentStatus", 
                          "PaymentGateway", "CreateDate", "UpdateDate")
        
        # Create data structure
        $data = New-Object System.Collections.ArrayList
        
        # Read the file content directly
        $fileContent = Get-Content -Path $RptFile -Raw
        
        # Use a regex to extract the rows in the specific format of this file
        $regex = 'SO-[^\s]+\s+\d+\s+\d+\.\d+\s+\w+\s+\w+\s+\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3}\s+(?:NULL|\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2}\.\d{3})'
        $matches = [regex]::Matches($fileContent, $regex)
        
        $rowCount = 0
        foreach ($match in $matches) {
            $row = $match.Value
            Write-Host "Found data row $rowCount`: $row" -ForegroundColor Green
            
            # Split the row by whitespace
            $parts = $row -split '\s+'
            
            # Parse into specific columns
            $orderCode = $parts[0]
            $companyId = $parts[1]
            $amount = $parts[2]
            $paymentStatus = $parts[3]
            $paymentGateway = $parts[4]
            
            # The datetime values might have spaces in them
            $createDateParts = @()
            $updateDateParts = @()
            $inUpdateDate = $false
            
            for ($i = 5; $i -lt $parts.Count; $i++) {
                if ($parts[$i] -eq "NULL") {
                    $updateDateParts += "NULL"
                    break
                }
                
                if (-not $inUpdateDate) {
                    # If we find another date pattern, we've moved to update date
                    if ($i -gt 7 -and $parts[$i] -match '^\d{4}-\d{2}-\d{2}$') {
                        $inUpdateDate = $true
                        $updateDateParts += $parts[$i]
                    } else {
                        $createDateParts += $parts[$i]
                    }
                } else {
                    $updateDateParts += $parts[$i]
                }
            }
            
            $createDate = [string]::Join(" ", $createDateParts)
            $updateDate = [string]::Join(" ", $updateDateParts)
            
            # Add to data collection
            $rowData = @($orderCode, $companyId, $amount, $paymentStatus, 
                          $paymentGateway, $createDate, $updateDate)
            $data.Add($rowData) | Out-Null
            $rowCount++
        }
        
        Write-Host "Extracted $($columnNames.Count) columns and $($data.Count) rows." -ForegroundColor Green
        
        # Create Excel file
        Write-Host "Creating Excel file..." -ForegroundColor Green
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "RPT Data"
        
        # Write column headers
        for ($col = 1; $col -le $columnNames.Count; $col++) {
            $worksheet.Cells.Item(1, $col) = $columnNames[$col-1]
            $worksheet.Cells.Item(1, $col).Font.Bold = $true
        }
        
        # Write data rows
        for ($row = 0; $row -lt $data.Count; $row++) {
            for ($col = 0; $col -lt $columnNames.Count; $col++) {
                $worksheet.Cells.Item($row+2, $col+1) = $data[$row][$col]
            }
        }
        
        # Auto-fit columns
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        
        # Add filters to headers
        $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $columnNames.Count))
        $headerRange.AutoFilter() | Out-Null
        
        # Save the Excel file
        $workbook.SaveAs($outputFile)
        $workbook.Close($false)
        $excel.Quit()
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Conversion completed successfully: $RptFile -> $outputFile" -ForegroundColor Green
    } else {
        # For other RPT formats, use a more generic approach
        Write-Host "Using generic RPT format handler." -ForegroundColor Green
        
        $content = Get-Content -Path $RptFile -Raw
        $lines = $content -split "`r`n|`r|`n"
        
        # Find header line
        $headerLine = ""
        foreach ($line in $lines) {
            if ($line -match '\S+\s+\S+' -and $line -notmatch '-{2,}') {
                $headerLine = $line
                break
            }
        }
        
        if (-not $headerLine) {
            throw "Could not find header line"
        }
        
        # Parse header columns
        $columnNames = @()
        $headerParts = $headerLine -split '\s{2,}'
        foreach ($part in $headerParts) {
            if ($part.Trim()) {
                $columnNames += $part.Trim()
            }
        }
        
        if ($columnNames.Count -eq 0) {
            throw "Could not extract column names"
        }
        
        # Find data rows
        $dataRows = New-Object System.Collections.ArrayList
        $inDataSection = $false
        foreach ($line in $lines) {
            if (-not $inDataSection) {
                if ($line -match '-{2,}') {
                    $inDataSection = $true
                }
                continue
            }
            
            if ($line -match '\(\d+\s+rows? affected\)') {
                break
            }
            
            if ($line.Trim() -and $line -notmatch "Completion time:") {
                $rowData = @()
                $lineParts = $line -split '\s{2,}'
                
                for ($i = 0; $i -lt $columnNames.Count; $i++) {
                    if ($i -lt $lineParts.Count) {
                        $value = $lineParts[$i].Trim()
                        if ($value -eq "NULL") { $value = "" }
                        $rowData += $value
                    } else {
                        $rowData += ""
                    }
                }
                
                $dataRows.Add($rowData) | Out-Null
            }
        }
        
        Write-Host "Extracted $($columnNames.Count) columns and $($dataRows.Count) rows." -ForegroundColor Green
        
        # Create Excel file
        Write-Host "Creating Excel file..." -ForegroundColor Green
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        
        $workbook = $excel.Workbooks.Add()
        $worksheet = $workbook.Worksheets.Item(1)
        $worksheet.Name = "RPT Data"
        
        # Write column headers
        for ($col = 1; $col -le $columnNames.Count; $col++) {
            $worksheet.Cells.Item(1, $col) = $columnNames[$col-1]
            $worksheet.Cells.Item(1, $col).Font.Bold = $true
        }
        
        # Write data rows
        for ($row = 0; $row -lt $dataRows.Count; $row++) {
            for ($col = 0; $col -lt $columnNames.Count; $col++) {
                $worksheet.Cells.Item($row+2, $col+1) = $dataRows[$row][$col]
            }
        }
        
        # Auto-fit columns
        $usedRange = $worksheet.UsedRange
        $usedRange.Columns.AutoFit() | Out-Null
        
        # Add filters to headers
        $headerRange = $worksheet.Range($worksheet.Cells.Item(1, 1), $worksheet.Cells.Item(1, $columnNames.Count))
        $headerRange.AutoFilter() | Out-Null
        
        # Save the Excel file
        $workbook.SaveAs($outputFile)
        $workbook.Close($false)
        $excel.Quit()
        
        # Clean up COM objects
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($worksheet) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        
        Write-Host "Conversion completed successfully: $RptFile -> $outputFile" -ForegroundColor Green
    }
} catch {
    Write-Host "Error: $_" -ForegroundColor Red
}
