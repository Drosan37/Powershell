#Requires -Version 5.1
<#
.SYNOPSIS
    Generates an HTML email report with a line chart (PNG, embedded as Base64)
    showing disk space usage per database over 10 days.

.DESCRIPTION
    1. Builds the chart as an SVG string (in memory).
    2. Renders the SVG to a PNG using System.Drawing (no external tools needed).
    3. Encodes the PNG as Base64 and embeds it directly in the HTML <img> tag.
    4. Sends the email via Send-MailMessage or falls back to Outlook COM.

.NOTES
    Requires: PowerShell 5.1 on Windows (.NET 4.x  System.Drawing is built-in).
#>
param(    
    [Parameter(ParameterSetName='Individual')]
    [string]$GroupName = '%'
)

Import-Module "$PSScriptRoot\..\Modules\mdl-SQLServer.psm1" -Force

# ---------------------------------------------------------------------------
# CONFIGURATION
# ---------------------------------------------------------------------------
$EmailConfig = @{
    From       = "monitoring@drosan37.it"
    To         = "alessandro.saracino@drosan37.it"
    Subject    = "Database Disk Space Report - $(Get-Date -Format 'yyyy-MM-dd')"
    SmtpServer = "smtp.drosan37.it"
    Port       = 25
    UseSsl     = $false
}

# ---------------------------------------------------------------------------
# STATIC DATA  (replace with real DB query results)
# ---------------------------------------------------------------------------
# Last 6 months  one label per month (current month last)
$DbName = "DBAdmin"
$strErrorLogFile = Join-Path $PSScriptRoot "\Logs\ErrorReportSpace.log"
# Set date now
$strDateNow = Get-Date -format "yyyyMMdd"

# Set errorlog file name
$strErrorLogFile = "{0}{1}_{2}" -f $strErrorLogFile.Substring(0,$strErrorLogFile.LastIndexOf('\')+1),$strDateNow,$strErrorLogFile.Substring($strErrorLogFile.LastIndexOf('\')+1);

# Define string for retrieve data about owner group
$strQueryOwnerGroup = ("
    SELECT  
          DatabaseName
        , InstanceName
        , GroupName
        , DestMail
  FROM [DBAdmin].[REP].[DestinationReport]
  WHERE GroupName LIKE '{0}'
" -f $GroupName)

# Check if group has been initialized (% means all group, so send mail to us)
if ($GroupName -ne '%')
{
    # Call function for retrieve data from DB
    $arrOwnerGroups = Get-DataFromIdera -Database $DbName -QueryToExec $strQueryOwnerGroup -LogPath $strErrorLogFile
    
    if ($arrOwnerGroups.Count -eq 0)
    {
        # No group has been found, so exit script
        Write-Host "No group identified"

        # Exit
        return
    }
    
    try
    {
        # Get dest mail (get only the first row)
        $strDestMail = $arrOwnerGroups.Get(0).DestMail
    }
    catch
    {
        # Get dest mail (it's only one row, so it's not an array)
        $strDestMail = $arrOwnerGroups["DestMail"]
    }  
    

    # Change dest mail
    $EmailConfig.To = $strDestMail
}

# Defile string for query insert, start with Truncate for remove old data
$strQueryServerList = ("
    SELECT 
        CONCAT(DatabaseName,'-',InstanceName) AS DatabaseName
        , AllDates
        , AllTotalAmounts
        ,  '#' + SUBSTRING(CONVERT(varchar(50), NEWID(), 2), 1, 6) AS Color
    FROM
    (
        SELECT  
			dk.DatabaseName,
	        dk.InstanceName,
            STRING_AGG(CONVERT(varchar(10), CONCAT(YEAR(dk.[QueryDate]),'-',RIGHT('0' + CAST(MONTH(dk.[QueryDate]) AS VARCHAR(2)),2)), 23), ',') 
                WITHIN GROUP (ORDER BY dk.QueryDate) AS AllDates,
            STRING_AGG(CAST(CAST(dk.[TotalSize]/1024/1024 AS DECIMAL(10,2)) AS VARCHAR(50)), ',') 
                WITHIN GROUP (ORDER BY dk.QueryDate) AS AllTotalAmounts
        FROM [DBAdmin].[REP].ReportDiskTopDatabases dk
		INNER JOIN [DBAdmin].[REP].[DestinationReport] dr
		ON dk.DatabaseName = dr.DatabaseName AND dk.InstanceName = dr.InstanceName
		WHERE dr.GroupName LIKE '{0}'
        GROUP BY dk.DatabaseName, dk.InstanceName
    ) tbl
    ORDER BY DatabaseName;  
;" -f $GroupName)

#Initialiye array
$Dates = @()
$Databases = @()

# Call function for retrieve data from DB
$arrTableRows = Get-DataFromIdera -Database $DbName -QueryToExec $strQueryServerList -LogPath $strErrorLogFile

# Cycle for add dates to array
for ($i = 5; $i -ge 0; $i--) {
    $Dates += (Get-Date).AddMonths(-$i).ToString("MM/yyyy")
}

# Cycle for each row in table output
foreach($row in $arrTableRows)
{
    # Add row in Database array
    $Databases += @{ Name = $row.DatabaseName; Color = $row.Color; Values = @($row.AllTotalAmounts.Split(',')) }
}

# ---------------------------------------------------------------------------
# STEP 1  BUILD SVG STRING
# ---------------------------------------------------------------------------
function Build-SVG {
    param(
        [array] $DbList,
        [array] $Labels,
        [int]   $Width      = 1024,
        [int]   $Height     = 468,
        [int]   $PadLeft    = 72,
        [int]   $PadRight   = 60,   # extra space so the last X-axis label is never clipped
        [int]   $PadTop     = 28,
        [int]   $PadBottom  = 50
        # Legend removed  colour is shown in the summary table instead
    )

    $plotW = $Width  - $PadLeft - $PadRight
    $plotH = $Height - $PadTop  - $PadBottom
    $n     = $Labels.Count

    # Y always starts at 0; top is max value + 5% headroom, rounded up to next integer
    $allVals = $DbList | ForEach-Object { $_.Values } | ForEach-Object { $_ }
    $dataMax = ($allVals | Measure-Object -Maximum).Maximum
    $minVal  = 0
    $maxVal  = [int][math]::Ceiling($dataMax * 1.05)
    $range   = $maxVal
    if ($range -eq 0) { $range = 1 }

    function xPos([int]$i) {
        return $PadLeft + [math]::Round(($i / ($n - 1)) * $plotW)
    }
    function yPos([double]$v) {
        return $PadTop + [math]::Round($plotH - ($v / $range) * $plotH)
    }

    $svg = [System.Text.StringBuilder]::new()
    [void]$svg.AppendLine("<svg xmlns='http://www.w3.org/2000/svg' width='$Width' height='$Height'>")

    # Background
    [void]$svg.AppendLine("  <rect width='$Width' height='$Height' fill='#FAFBFD'/>")

    # Grid lines + Y labels  5 evenly-spaced steps from 0 to maxVal, integer labels
    $steps = 5
    for ($s = 0; $s -le $steps; $s++) {
        $frac = $s / $steps
        $gVal = [int][math]::Round($frac * $maxVal)
        $gy   = $PadTop + [math]::Round($plotH - $frac * $plotH)
        $gc   = if ($s -eq 0) { "#AAAAAA" } else { "#E0E0E0" }
        [void]$svg.AppendLine("  <line x1='$PadLeft' y1='$gy' x2='$($PadLeft + $plotW)' y2='$gy' stroke='$gc' stroke-width='1'/>")
        [void]$svg.AppendLine("  <text x='$($PadLeft - 6)' y='$($gy + 4)' text-anchor='end' font-size='11' font-family='Arial' fill='#666666'>${gVal} TB</text>")
    }

    # X axis labels
    # Last label uses text-anchor='end' so it stays within the right padding
    for ($i = 0; $i -lt $n; $i++) {
        $lx     = xPos $i
        $ly     = $PadTop + $plotH + 18
        $anchor = if ($i -eq ($n - 1)) { "end" } else { "middle" }
        [void]$svg.AppendLine("  <text x='$lx' y='$ly' text-anchor='$anchor' font-size='11' font-family='Arial' fill='#666666'>$($Labels[$i])</text>")
    }

    # Axes
    $xAxisY = $PadTop + $plotH
    [void]$svg.AppendLine("  <line x1='$PadLeft' y1='$xAxisY' x2='$($PadLeft + $plotW)' y2='$xAxisY' stroke='#999999' stroke-width='1'/>")
    [void]$svg.AppendLine("  <line x1='$PadLeft' y1='$PadTop' x2='$PadLeft' y2='$xAxisY' stroke='#999999' stroke-width='1'/>")

    # Lines + dots per database  (no legend  see summary table)
    foreach ($db in $DbList) {
        $col  = $db.Color
        $vals = $db.Values

        # polyline
        $pts = @()
        for ($i = 0; $i -lt $n; $i++) {
            $pts += "$((xPos $i)),$((yPos $vals[$i]))"
        }
        $pointsStr = $pts -join " "
        [void]$svg.AppendLine("  <polyline points='$pointsStr' fill='none' stroke='$col' stroke-width='2.5' stroke-linejoin='round' stroke-linecap='round'/>")

        # dots
        for ($i = 0; $i -lt $n; $i++) {
            $cx = xPos $i
            $cy = yPos $vals[$i]
            [void]$svg.AppendLine("  <circle cx='$cx' cy='$cy' r='4' fill='$col' stroke='#ffffff' stroke-width='2'/>")
        }
    }

    [void]$svg.AppendLine("</svg>")
    return $svg.ToString()
}

# ---------------------------------------------------------------------------
# STEP 2  RENDER SVG to PNG VIA System.Drawing (GDI+)
# No external binaries required  pure .NET
# ---------------------------------------------------------------------------
function ConvertSVG-ToPngBase64 {
    param(
        [string] $SvgString,
        [int]    $Width  = 700,
        [int]    $Height = 320
    )

    Add-Type -AssemblyName System.Drawing

    # Helper: "#RRGGBB" -> System.Drawing.Color
    function HexToColor([string]$hex) {
        $hex = $hex.TrimStart('#')
        return [System.Drawing.Color]::FromArgb(
            [Convert]::ToInt32($hex.Substring(0,2),16),
            [Convert]::ToInt32($hex.Substring(2,2),16),
            [Convert]::ToInt32($hex.Substring(4,2),16)
        )
    }

    $bmp = New-Object System.Drawing.Bitmap($Width, $Height)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    $g.SmoothingMode      = [System.Drawing.Drawing2D.SmoothingMode]::AntiAlias
    $g.TextRenderingHint  = [System.Drawing.Text.TextRenderingHint]::AntiAlias
    $g.Clear([System.Drawing.Color]::FromArgb(250, 251, 253))

    # Parse SVG XML
    $xml = [xml]$SvgString

    foreach ($node in $xml.svg.ChildNodes) {
        switch ($node.LocalName) {

            # ---- rect ----
            "rect" {
                if ($node.fill -and $node.fill -ne "none") {
                    $c  = HexToColor $node.fill
                    $br = New-Object System.Drawing.SolidBrush($c)
                    $g.FillRectangle($br,
                        [float]$node.x, [float]$node.y,
                        [float]$node.width, [float]$node.height)
                    $br.Dispose()
                }
            }

            # ---- line ----
            "line" {
                $sc = $node.stroke
                $sw = if ($node.'stroke-width') { [float]$node.'stroke-width' } else { 1.0 }
                if ($sc -and $sc -ne "none") {
                    $pen = New-Object System.Drawing.Pen((HexToColor $sc), $sw)
                    $g.DrawLine($pen,
                        [float]$node.x1, [float]$node.y1,
                        [float]$node.x2, [float]$node.y2)
                    $pen.Dispose()
                }
            }

            # ---- polyline ----
            "polyline" {
                $sc = $node.stroke
                $sw = if ($node.'stroke-width') { [float]$node.'stroke-width' } else { 1.0 }
                if ($sc -and $sc -ne "none") {
                    $pen           = New-Object System.Drawing.Pen((HexToColor $sc), $sw)
                    $pen.LineJoin  = [System.Drawing.Drawing2D.LineJoin]::Round
                    $pen.StartCap  = [System.Drawing.Drawing2D.LineCap]::Round
                    $pen.EndCap    = [System.Drawing.Drawing2D.LineCap]::Round

                    $ptArr = $node.points.Trim() -split '\s+'
                    $ptList = New-Object System.Collections.Generic.List[System.Drawing.PointF]
                    foreach ($pt in $ptArr) {
                        $xy = $pt -split ','
                        if ($xy.Count -eq 2) {
                            $ptList.Add([System.Drawing.PointF]::new([float]$xy[0],[float]$xy[1]))
                        }
                    }
                    if ($ptList.Count -ge 2) {
                        $g.DrawLines($pen, $ptList.ToArray())
                    }
                    $pen.Dispose()
                }
            }

            # ---- circle ----
            "circle" {
                $cx = [float]$node.cx
                $cy = [float]$node.cy
                $r  = [float]$node.r

                if ($node.fill -and $node.fill -ne "none") {
                    $br = New-Object System.Drawing.SolidBrush((HexToColor $node.fill))
                    $g.FillEllipse($br, $cx - $r, $cy - $r, $r*2, $r*2)
                    $br.Dispose()
                }
                if ($node.stroke -and $node.stroke -ne "none") {
                    $sw  = if ($node.'stroke-width') { [float]$node.'stroke-width' } else { 1.0 }
                    $pen = New-Object System.Drawing.Pen((HexToColor $node.stroke), $sw)
                    $g.DrawEllipse($pen, $cx - $r, $cy - $r, $r*2, $r*2)
                    $pen.Dispose()
                }
            }

            # ---- text ----
            "text" {
                $fontSize   = if ($node.'font-size') { [float]$node.'font-size' } else { 12.0 }
                $fillColor  = if ($node.fill)        { $node.fill }               else { "#000000" }
                $font       = New-Object System.Drawing.Font("Arial", $fontSize)
                $br         = New-Object System.Drawing.SolidBrush((HexToColor $fillColor))
                $tx         = [float]$node.x
                $ty         = [float]$node.y - $fontSize   # SVG baseline -> top-left
                $textVal    = $node.InnerText

                switch ($node.'text-anchor') {
                    "end"    { $sz = $g.MeasureString($textVal,$font); $tx = $tx - $sz.Width }
                    "middle" { $sz = $g.MeasureString($textVal,$font); $tx = $tx - ($sz.Width / 2) }
                }

                $g.DrawString($textVal, $font, $br, $tx, $ty)
                $font.Dispose()
                $br.Dispose()
            }
        }
    }

    $g.Dispose()

    # Save to MemoryStream → Base64
    $ms = New-Object System.IO.MemoryStream
    $bmp.Save($ms, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    $base64 = [Convert]::ToBase64String($ms.ToArray())
    $ms.Dispose()
    return $base64
}

# ---------------------------------------------------------------------------
# STEP 3  SUMMARY TABLE
# Columns: Colour | Database | <Month1> | <Month2>  <MonthN> | Avg Trend
# ---------------------------------------------------------------------------
function Build-SummaryTable {
    param([array]$DbList, [array]$Labels)

    $n       = $Labels.Count
    $lastIdx = $n - 1

    # ---- Header row ----
    $thStyle   = "padding:8px 12px;font-family:Calibri,Arial,sans-serif;font-size:12px;color:#555;font-weight:600;border-bottom:2px solid #E0E0E0;white-space:nowrap;"
    $headerRow = @"
      <tr style="background:#F7F9FC;">
        <th style="$thStyle text-align:center;width:36px;">Colour</th>
        <th style="$thStyle text-align:left;">Database</th>
"@
    foreach ($lbl in $Labels) {
        $headerRow += "        <th style=`"$thStyle text-align:right;`">$lbl</th>`n"
    }
    $headerRow += "        <th style=`"$thStyle text-align:right;`">Avg&nbsp;Trend</th>`n"
    $headerRow += "      </tr>"

    # ---- Data rows ----
    $tdBase  = "padding:8px 12px;font-family:Calibri,Arial,sans-serif;font-size:13px;white-space:nowrap;"
    $rows    = ""

    foreach ($db in $DbList) {
        $vals = $db.Values

        # Monthly deltas (month-over-month change), starting from index 1
        $deltas = @()
        for ($i = 1; $i -lt $n; $i++) {
            $deltas += [math]::Round($vals[$i] - $vals[$i-1], 2)
        }

        # Average trend = mean of all month-over-month deltas
        $avgTrend  = [math]::Round(($deltas | Measure-Object -Average).Average, 2)
        $avgArrow  = if ($avgTrend -gt 0) { "&#9650;" } else { "&#9660;" }
        $avgColor  = if ($avgTrend -gt 0) { "#E74C3C" } else { "#27AE60" }
        $avgSign   = if ($avgTrend -gt 0) { "+" }       else { "" }

        # Build month value cells  highlight last month bold
        $monthCells = ""
        for ($i = 0; $i -lt $n; $i++) {
            $v       = [math]::Round($vals[$i], 2)
            $weight  = if ($i -eq $lastIdx) { "font-weight:700;" } else { "" }
            $monthCells += "          <td style=`"$tdBase $weight color:#333;text-align:right;`">${v}&nbsp;TB</td>`n"
        }

        $rows += @"
        <tr style="border-bottom:1px solid #F0F0F0;">
          <td style="$tdBase text-align:center;width:36px;">
            <table cellpadding="0" cellspacing="0" border="0" style="margin:0 auto;"><tr>
              <td style="width:22px;height:10px;background:$($db.Color);border-radius:3px;font-size:0;">&nbsp;</td>
            </tr></table>
          </td>
          <td style="$tdBase color:#333;font-weight:600;">$($db.Name)</td>
$monthCells          <td style="$tdBase color:${avgColor};text-align:right;">${avgArrow}&nbsp;${avgSign}${avgTrend}&nbsp;TB</td>
        </tr>
"@
    }

    return @"
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="border-collapse:collapse;border:1px solid #E0E0E0;">
$headerRow
$rows
    </table>
"@
}

# ---------------------------------------------------------------------------
# BUILD EVERYTHING
# ---------------------------------------------------------------------------
Write-Host "[1/4] Building SVG chart..." -ForegroundColor Cyan
$svgString = Build-SVG -DbList $Databases -Labels $Dates -Width 1024 -Height 468

Write-Host "[2/4] Rendering SVG -> PNG (System.Drawing)..." -ForegroundColor Cyan
$pngBase64 = ConvertSVG-ToPngBase64 -SvgString $svgString -Width 1024 -Height 468

Write-Host "[3/4] Assembling HTML email..." -ForegroundColor Cyan
$summaryHtml = Build-SummaryTable -DbList $Databases -Labels $Dates
$reportDate  = Get-Date -Format "dddd, MMMM d, yyyy"
$genTime     = Get-Date -Format "HH:mm"

# ---------------------------------------------------------------------------
# STEP 4  FULL HTML EMAIL
# ---------------------------------------------------------------------------
$htmlBody = @"
<!DOCTYPE html>
<html>
<head>
  <meta http-equiv="Content-Type" content="text/html; charset=UTF-8">
  <title>Database Disk Space Report</title>
</head>
<body style="margin:0;padding:0;background:#F4F6F9;">

  <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background:#F4F6F9;padding:24px 0;">
    <tr><td align="center">

      <table width="1088" cellpadding="0" cellspacing="0" border="0" style="background:#FFFFFF;">

        <!-- HEADER -->
        <tr>
          <td style="background:#002149;padding:24px 32px;">
            <p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:20px;font-weight:700;color:#FFFFFF;">
              Database Disk Space Report
            </p>
            <p style="margin:4px 0 0;font-family:Calibri,Arial,sans-serif;font-size:13px;color:#CCE4F7;">
              $reportDate &nbsp;&bull;&nbsp; Generated at $genTime
            </p>
          </td>
        </tr>

        <!-- INTRO -->
        <tr>
          <td style="padding:24px 32px 12px;">
            <p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:14px;color:#555;line-height:1.6;">
              The chart below shows total allocated disk space (TB) for each monitored
              database over the last <strong>6 months</strong>.
            </p>
          </td>
        </tr>

        <!-- CHART  PNG embedded as Base64 (Outlook safe) -->
        <tr>
          <td style="padding:8px 32px 4px;" align="center">
            <table cellpadding="0" cellspacing="0" border="0"
                   style="background:#FAFBFD;border:1px solid #E8ECF0;">
              <tr>
                <td style="padding:16px;" align="center">
                  <img src="data:image/png;base64,$pngBase64"
                       width="1024" height="468" border="0"
                       alt="Database Disk Space Chart"
                       style="display:block;" />
                </td>
              </tr>
              <tr>
                <td style="padding:0 16px 10px;" align="center">
                  <p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11px;color:#999;">
                    Figure 1  Total disk space (TB) per database  last 6 months
                  </p>
                </td>
              </tr>
            </table>
          </td>
        </tr>

        <!-- SUMMARY TABLE -->
        <tr>
          <td style="padding:20px 32px 28px;">
            <p style="margin:0 0 10px;font-family:Calibri,Arial,sans-serif;font-size:15px;font-weight:700;color:#333;">
              Current Snapshot
            </p>
            $summaryHtml
          </td>
        </tr>

        <!-- DIVIDER -->
        <tr>
          <td style="padding:0 32px;">
            <table width="100%" cellpadding="0" cellspacing="0" border="0">
              <tr><td style="height:1px;background:#EEEEEE;"></td></tr>
            </table>
          </td>
        </tr>

        <!-- FOOTER -->
        <tr>
          <td style="padding:16px 32px;">
            <p style="margin:0;font-family:Calibri,Arial,sans-serif;font-size:11px;color:#AAAAAA;line-height:1.5;">
              Automated report  Database Monitoring System.<br>
              Do not reply. Questions? <a href="mailto:alessandro.saracino@drosan37.it" style="color:#0078D4;">Database Administrator Team</a>
            </p>
          </td>
        </tr>

      </table>
    </td></tr>
  </table>

</body>
</html>
"@

# ---------------------------------------------------------------------------
# SAVE TO ReportHTML SUBFOLDER
# ---------------------------------------------------------------------------

# Ensure the ReportHTML subfolder exists next to the script
$reportFolder = Join-Path $PSScriptRoot "ReportHTML"
if (-not (Test-Path $reportFolder)) {
    New-Item -ItemType Directory -Path $reportFolder | Out-Null
    Write-Host "[INFO] Created folder: $reportFolder" -ForegroundColor DarkCyan
}

$htmlOutputPath = Join-Path $reportFolder "DiskSpaceReport_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$htmlBody | Out-File -FilePath $htmlOutputPath -Encoding UTF8
Write-Host "[3/4] HTML report saved: $htmlOutputPath" -ForegroundColor Cyan

# ---------------------------------------------------------------------------
# CLEANUP  keep only reports generated TODAY, delete all older ones
# ---------------------------------------------------------------------------
function Invoke-ReportCleanup {
    param(
        [string] $Folder,
        [string] $Pattern = "DiskSpaceReport_*.html"
    )

    $today      = (Get-Date).Date          # midnight of today  (e.g. 2026-03-17 00:00:00)
    $todayStamp = Get-Date -Format "yyyyMMdd"   # used to match filename prefix

    Write-Host "[CLEANUP] Scanning '$Folder' for old reports..." -ForegroundColor DarkCyan

    $allReports = Get-ChildItem -Path $Folder -Filter $Pattern -File -ErrorAction SilentlyContinue

    if (-not $allReports) {
        Write-Host "[CLEANUP] No report files found  nothing to clean." -ForegroundColor DarkGray
        return
    }

    $deleted = 0
    $kept    = 0

    foreach ($file in $allReports) {
        # Primary check: file's LastWriteTime date (reliable for files written by this script)
        $fileDate = $file.LastWriteTime.Date

        # Secondary check: date embedded in the filename  "DiskSpaceReport_20260317_143022.html"
        $fileNameDate = $null
        if ($file.BaseName -match 'DiskSpaceReport_(\d{8})_') {
            try { $fileNameDate = [datetime]::ParseExact($Matches[1], "yyyyMMdd", $null) }
            catch { }
        }

        # A file is "today's" only when BOTH checks agree it belongs to today
        $isToday = ($fileDate -eq $today) -and
                   ((-not $fileNameDate) -or ($fileNameDate -eq $today))

        if ($isToday) {
            $kept++
            Write-Host "  [KEEP]   $($file.Name)" -ForegroundColor DarkGray
        }
        else {
            try {
                Remove-Item -Path $file.FullName -Force -ErrorAction Stop
                $deleted++
                Write-Host "  [DELETE] $($file.Name)  (date: $($fileDate.ToString('yyyy-MM-dd')))" -ForegroundColor Yellow
            }
            catch {
                Write-Warning "  [WARN] Could not delete '$($file.Name)': $_"
            }
        }
    }

    Write-Host "[CLEANUP] Done kept: $kept  |  deleted: $deleted" -ForegroundColor DarkCyan
}

Invoke-ReportCleanup -Folder $reportFolder

# ---------------------------------------------------------------------------
# SEND
# ---------------------------------------------------------------------------
Write-Host "[4/4] Sending email..." -ForegroundColor Cyan
try {
    $mailParams = @{
        From       = $EmailConfig.From
        To         = $EmailConfig.To
        Subject    = $EmailConfig.Subject
        Body       = $htmlBody
        BodyAsHtml = $true
        SmtpServer = $EmailConfig.SmtpServer
        Port       = $EmailConfig.Port
        Encoding   = [System.Text.Encoding]::UTF8
    }
    if ($EmailConfig.UseSsl) { $mailParams['UseSsl'] = $true }

    #TEST - Disable send mail
    Send-MailMessage @mailParams
    Write-Host "[OK] Email sent via Send-MailMessage." -ForegroundColor Green
}
catch {
    Write-Warning "[WARN] Send-MailMessage failed: $_"
    Write-Host "      Falling back to Outlook COM..." -ForegroundColor Yellow
    try {
        $outlook       = New-Object -ComObject Outlook.Application
        $mail          = $outlook.CreateItem(0)
        $mail.To       = $EmailConfig.To
        $mail.Subject  = $EmailConfig.Subject
        $mail.HTMLBody = $htmlBody
        $mail.Send()
        Write-Host "[OK] Email sent via Outlook COM." -ForegroundColor Green
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail)    | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) | Out-Null
    }
    catch {
        Write-Error "[ERROR] Both methods failed: $_"
    }
}
