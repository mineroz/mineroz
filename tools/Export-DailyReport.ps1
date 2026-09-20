<#
.SYNOPSIS
    Exports the Daily_Report sheet of an MPS workbook to PDF with Excel, optionally into an Outlook mail.
.DESCRIPTION
    Uses Excel COM automation (the workbook contains no macros). The script opens the workbook
    read-only, writes the report date into Daily_Report!C5, lets Excel calculate, exports the
    Daily_Report sheet to

        <OutDir>\Mako_Daily_Report_yyyy-MM-dd.pdf

    and closes Excel without saving, so the workbook is left exactly as it was. With -Email an
    Outlook message with the PDF attached is displayed for review; nothing is sent automatically.
    Requires Microsoft Excel (and classic Outlook for -Email) on this PC.
.PARAMETER Workbook
    Path to the MPS workbook. Default: <Root>\04_MPS\MPS_<year of Date>.xlsx.
.PARAMETER Date
    Report date, for example 2026-03-15. Default: yesterday.
.PARAMETER OutDir
    Folder for the PDF. Default: <workbook folder>\Reports (created when missing).
.PARAMETER Root
    Working folder root used for the default workbook path. Default C:\MakoPS.
.PARAMETER Sheet
    Report sheet name. Default Daily_Report.
.PARAMETER DateCell
    Cell holding the report date on that sheet. Default C5.
.PARAMETER Email
    Create an Outlook message with the PDF attached and display it. The user reviews and sends it.
.PARAMETER To
    Recipients for -Email. One or more addresses.
.PARAMETER Cc
    Copy recipients for -Email.
.PARAMETER Subject
    Mail subject. Default "Mako Daily Production Report <date>".
.PARAMETER Open
    Open the PDF when done.
.EXAMPLE
    .\Export-DailyReport.ps1
    .\Export-DailyReport.ps1 -Workbook C:\MakoPS\04_MPS\MPS_2026_DEMO.xlsx -Date 2026-03-15 -Open
    .\Export-DailyReport.ps1 -Date 2026-03-15 -Email -To gm@example.com,ops@example.com
#>
[CmdletBinding()]
param(
    [string]$Workbook = "",
    [datetime]$Date = (Get-Date).Date.AddDays(-1),
    [string]$OutDir = "",
    [string]$Root = "C:\MakoPS",
    [string]$Sheet = "Daily_Report",
    [string]$DateCell = "C5",
    [switch]$Email,
    [string[]]$To = @(),
    [string[]]$Cc = @(),
    [string]$Subject = "",
    [switch]$Open
)

# No native commands run here, so Stop is safe and makes every COM failure land in the catch/finally blocks.
$ErrorActionPreference = "Stop"

$Date = $Date.Date
if ([string]::IsNullOrWhiteSpace($Workbook)) { $Workbook = Join-Path $Root ("04_MPS\MPS_{0}.xlsx" -f $Date.Year) }
if (-not (Test-Path -LiteralPath $Workbook)) { throw "Workbook not found: $Workbook  (run Build-MPS.ps1 first, or pass -Workbook)" }
$Workbook = (Resolve-Path -LiteralPath $Workbook).ProviderPath
if ([string]::IsNullOrWhiteSpace($OutDir)) { $OutDir = Join-Path (Split-Path -Parent $Workbook) "Reports" }
if (-not (Test-Path -LiteralPath $OutDir)) { New-Item -ItemType Directory -Force -Path $OutDir | Out-Null }
$OutDir = (Resolve-Path -LiteralPath $OutDir).ProviderPath
$pdf = Join-Path $OutDir ("Mako_Daily_Report_{0:yyyy-MM-dd}.pdf" -f $Date)
if ($Email -and $To.Count -eq 0) { throw "-Email needs at least one recipient, for example: -To name@example.com" }

function Remove-ComRef($obj) {
    # Releases the runtime callable wrapper so the Excel / Outlook process can exit.
    if ($null -ne $obj) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($obj) } catch { }
    }
}

# ---------------------------------------------------------------- Excel: set the date, calculate, export
$excel = $null; $books = $null; $wb = $null; $sheets = $null; $ws = $null; $cell = $null
try {
    try { $excel = New-Object -ComObject Excel.Application }
    catch {
        throw ("Microsoft Excel is not installed on this computer, or COM automation is blocked, so the PDF cannot be produced here. " +
               "Run this script on a PC with Excel, or open the workbook, set {0}!{1} to the report date and use File > Export > Create PDF." -f $Sheet, $DateCell)
    }
    $excel.Visible = $false
    $excel.DisplayAlerts = $false
    $excel.ScreenUpdating = $false
    $excel.AskToUpdateLinks = $false
    $excel.EnableEvents = $false
    try { $excel.AutomationSecurity = 3 } catch { }   # msoAutomationSecurityForceDisable: the workbook has no macros

    Write-Host "Opening $Workbook (read-only)"
    $books = $excel.Workbooks
    $wb = $books.Open($Workbook, 0, $true)            # UpdateLinks = 0, ReadOnly = true
    $sheets = $wb.Worksheets
    try { $ws = $sheets.Item($Sheet) } catch { throw "Sheet '$Sheet' not found in $Workbook" }

    # C5 is unlocked in the protected sheet, so it can be written without unprotecting anything.
    $cell = $ws.Range($DateCell)
    $cell.Value2 = $Date.ToOADate()

    $excel.CalculateFull()
    $deadline = (Get-Date).AddMinutes(5)
    while ($excel.CalculationState -ne 0 -and (Get-Date) -lt $deadline) { Start-Sleep -Milliseconds 250 }   # 0 = xlDone

    # Warn when the requested date lies beyond the entered data (the report would be blank).
    $nm = $null; $rng = $null
    try {
        $nm = $wb.Names.Item("cfg_LastDataDate")
        $rng = $nm.RefersToRange
        $last = $rng.Value2
        if ($last -is [double] -and $last -gt 0) {
            $lastDate = [datetime]::FromOADate($last)
            if ($Date -gt $lastDate) {
                Write-Warning ("Last entered data is {0:yyyy-MM-dd}; the report for {1:yyyy-MM-dd} will show no figures." -f $lastDate, $Date)
            }
        }
    } catch { }
    finally { Remove-ComRef $rng; Remove-ComRef $nm }

    if (Test-Path -LiteralPath $pdf) { Remove-Item -LiteralPath $pdf -Force }
    Write-Host ("Exporting {0} for {1:ddd dd-MMM-yyyy} to PDF" -f $Sheet, $Date)
    # ExportAsFixedFormat(Type, Filename, Quality, IncludeDocProperties, IgnorePrintAreas): xlTypePDF = 0, xlQualityStandard = 0
    $ws.ExportAsFixedFormat(0, $pdf, 0, $true, $false)
    if (-not (Test-Path -LiteralPath $pdf)) { throw "Excel did not write $pdf" }
    Write-Host "PDF: $pdf"
}
finally {
    if ($null -ne $wb) { try { $wb.Close($false) } catch { } }          # never save
    if ($null -ne $excel) { try { $excel.Quit() } catch { } }
    foreach ($o in @($cell, $ws, $sheets, $wb, $books, $excel)) { Remove-ComRef $o }
    $cell = $null; $ws = $null; $sheets = $null; $wb = $null; $books = $null; $excel = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers(); [GC]::Collect()
}

# ---------------------------------------------------------------- Outlook: draft the mail (displayed, not sent)
if ($Email) {
    $ol = $null; $mail = $null; $att = $null
    try {
        try { $ol = New-Object -ComObject Outlook.Application }
        catch { throw "Classic Microsoft Outlook is not available on this computer (the new Outlook has no COM automation). The PDF is at $pdf; attach it to a mail manually." }
        if ([string]::IsNullOrWhiteSpace($Subject)) { $Subject = ("Mako Daily Production Report {0:ddd dd-MMM-yyyy}" -f $Date) }
        $mail = $ol.CreateItem(0)                     # olMailItem
        $mail.To = ($To -join "; ")
        if ($Cc.Count -gt 0) { $mail.CC = ($Cc -join "; ") }
        $mail.Subject = $Subject
        $mail.Body = ("Please find attached the Mako daily production report for {0:dddd dd MMMM yyyy}.`r`n`r`nSource workbook: {1}`r`n" -f $Date, (Split-Path -Leaf $Workbook))
        $att = $mail.Attachments
        [void]$att.Add($pdf)
        $mail.Display()
        Write-Host "Outlook message opened for review. Check the recipients and send it yourself."
    }
    finally {
        foreach ($o in @($att, $mail, $ol)) { Remove-ComRef $o }
        $att = $null; $mail = $null; $ol = $null
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }
}

if ($Open) { Start-Process -FilePath $pdf }
