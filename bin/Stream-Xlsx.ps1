#!/usr/bin/env pwsh

Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, "netstandard.dll"))
Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, "Sylvan.Data.Excel.dll"))
Add-Type -Path ([System.IO.Path]::Combine($PSScriptRoot, "Sylvan.Data.Csv.dll"))

$stdin = [System.Console]::OpenStandardInput()
$stdout = [System.Console]::Out

$reader = [Sylvan.Data.Excel.ExcelDataReader]::Create($stdin, [Sylvan.Data.Excel.ExcelWorkbookType]::ExcelXml)
$writer = [Sylvan.Data.Csv.CsvDataWriter]::Create($stdout)

# We can iterate over worksheets with ExcelDataReader.NextResult() and
# retrieve the worksheet name with ExcelDataReader.WorksheetName, but we
# can't modify the source field midstream. Only the first worksheet's
# header row will be recognized by INDEX_EXTRACTIONS = csv.

#do {
#    $writer.Write($reader) | Out-Null
#} while ($reader.NextResult())

$writer.Write($reader) | Out-Null
