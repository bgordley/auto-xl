param(
    [string]$c, # Command
    [string]$p  # Path of json file
)

. "U:\Software Engineering Team\Reporting\Scripts\Tools\EnvVars.ps1"
. "U:\Software Engineering Team\Reporting\Scripts\Tools\GUITools.ps1"

function TestFile([string]$filePath) {
    # Test the given file's integrity then return test result
    $isValid = $false

    if(Test-Path $filePath) {
        if($(".xlsx.xlsm").Contains([System.IO.Path]::GetExtension($filePath))) {
            try { 
                [IO.File]::OpenWrite($filePath).close() 
                $isValid = $true
            } catch { 
                $isValid = $false
            }
        }
    }

    return $isValid
}

switch ($c.ToUpper()) {
    'INIT' {
        $excel = @{
            path = 'C:\example\example.xlsm';
            output = 'C:\example\prod\';
            macro = 'Example';
        }
        $init = @{}
        $init.excelTemplates = @($excel)
        $init.preRunCommands = @('C:\example\')
        $init.postRunCommands = @('C:\example\')
        $init | ConvertTo-Json | Out-File $p
        exit
    }

    'RUN' {
        $st = if(Test-Path $p) { $(Get-Content $p | ConvertFrom-Json) } else { Write-Debug "Settings file $p could not be found"; exit }
        $currDate = $(Get-Date -Format "yyyyMMdd")
        
        # Test file integrity
        $st.excelTemplates | ForEach-Object { 
            $result = $(TestFile($_.path))
            Write-Debug "File $($_.path) test passed: $result"

            # On error, exit script
            if($result -eq $false) { 
                Write-Debug "File integrity check failed"
                exit
            }
        }

        # Run all pre-run commands
        $st.preRunCommands | ForEach-Object {
            try{ if($_ -ne $null -and $_ -ne "" ) { Invoke-Expression $_ }}
            catch{ Write-Debug "Failed to run command $_"; exit }
        }

        # Initialize Excel object
        Add-Type -AssemblyName Microsoft.Office.Interop.Excel
        $xlFixedFormat = [Microsoft.Office.Interop.Excel.XlFileFormat]::xlWorkbookDefault
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $st.excelTemplates | ForEach-Object {
            # Output parameters for review in case of error
            Write-Debug "Excel Template Params: $_"

            # Try to open the workbook
            try{ $workbook = $excel.Workbooks.Open($_.path); Write-Debug "Successfully opened $($_.path)"}
            catch{ Write-Debug "Failed to open $_"; exit }

            # Try to run the provided macro if not blank
            try{ if($_.macro.Trim() -ne "") { $excel.Run($_.macro); Write-Debug "Macro $($_.macro) has run successfully" }}
            catch{ Write-Debug "Failed to run macro $($_.macro)"; exit }

            # Save the report to the provided output location
            $outFile = [System.IO.Path]::combine($_.output, [System.IO.Path]::GetFileNameWithoutExtension($_.path) + $currDate + ".xlsx")
            try{ $workbook.SaveAs($outFile, $xlFixedFormat); Write-Debug "Successfully saved as $outFile" }
            catch{ Write-Debug "Failed to save as $outFile"; exit }

            $excel.Quit()
        }

        # Run all post-run commands
        $st.postRunCommands | ForEach-Object { 
            try{ if($_ -ne $null) { Invoke-Expression $_ }}
            catch{ Write-Debug "Failed to run command $_"; exit }
        }
    }
}