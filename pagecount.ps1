<#
.SYNOPSIS
    PowerShell script to collect the page count of pdf documents on multiple network sources
.DESCRIPTION
    This script returns an Excel spreadsheet with a list of the pdfs in the specified folder and the page count of each pdf.
.EXAMPLE
    pagecount.ps1 -filePath "C:\Users\User\Desktop"
#>
Param(
    [string]$filePath = (Get-Location).Path
)

function Write-Log {
    param (
        [string]$message,
        [switch]$logOnly = $false
    )
    $now = Get-Date  
    $log = $now.ToString() + " - $message"
    if (-not($logOnly)) {
        Write-Host $log
    }
    $log | Out-File -filepath "$PSScriptRoot\log.txt" -Append
}

Write-Log "Processing $filePath"
$pdfs = Get-ChildItem $filePath -Filter *.pdf -Recurse -Force
$total = $pdfs.Count
Write-Log "Found $total pdf files."
$EA = $ErrorActionPreference
$WA = $WarningPreference
$ErroractionPreference = "SilentlyContinue"
$WarningPreference = "SilentlyContinue"
$count = $pdfs | ForEach-Object -Parallel {
    if($file = Get-PDF $_.FullName -WarningVariable warn){
        [PSCustomObject]@{
            Name  = $_.Name
            Pages = $file.GetNumberOfPages()
            Notes = ""
        } 
    }
    else {
        [PSCustomObject]@{
            Name  = $_.Name
            Pages = 0
            Notes = $warn.message
        } 
    }
}
Write-Log "Finished processing $filePath"
$ErrorActionPreference = $EA
$WarningPreference = $WA
$count | Export-Excel "$PSScriptRoot\count.xlsx" -AutoSize
$file = Open-ExcelPackage "$PSScriptRoot\count.xlsx"
$countCell = "C" + ($file.Sheet1.Dimension.Rows + 1).tostring()
$totalCell = "B" + ($file.Sheet1.Dimension.Rows + 1).tostring()
$file.Sheet1.cells[$countCell].Formula = "=SUM(B:B)"
$file.Workbook.Worksheets['Sheet1'].Cells[$countCell] | Set-Format -HorizontalAlignment Left
$file.Workbook.Worksheets['Sheet1'].Cells[$totalCell].Value = "Total"
$file.Workbook.Worksheets['Sheet1'].Cells[$totalCell] | Set-Format -HorizontalAlignment Right
Close-ExcelPackage $file
