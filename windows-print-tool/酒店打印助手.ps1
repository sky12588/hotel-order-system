Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"
$BaseDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$ConfigPath = Join-Path $BaseDir "打印助手配置.json"
$DownloadDir = Join-Path $BaseDir "已接收打印任务"
New-Item -ItemType Directory -Force -Path $DownloadDir | Out-Null

function Load-Config {
    if (!(Test-Path $ConfigPath)) {
        throw "找不到配置文件：$ConfigPath"
    }
    return Get-Content $ConfigPath -Encoding UTF8 | ConvertFrom-Json
}

function Get-InstalledPrinters {
    try {
        return @(Get-CimInstance Win32_Printer | Sort-Object Name | ForEach-Object { $_.Name })
    } catch {
        return @()
    }
}

function Resolve-Printer($preferred, $kind) {
    $printers = Get-InstalledPrinters
    if ($preferred -and ($printers -contains $preferred)) {
        return $preferred
    }
    if ($kind -eq "sales") {
        $match = $printers | Where-Object { $_ -match "EPSON|LQ|610|针式" } | Select-Object -First 1
    } else {
        $match = $printers | Where-Object { $_ -match "HP|Laser|激光" } | Select-Object -First 1
    }
    if ($match) { return $match }
    return $printers | Select-Object -First 1
}

function Get-Printer-Info($printerName) {
    if (!$printerName) { return $null }
    try {
        return Get-CimInstance Win32_Printer | Where-Object { $_.Name -eq $printerName } | Select-Object -First 1
    } catch {
        return $null
    }
}

function Get-Printer-StatusText($printerName) {
    $printer = Get-Printer-Info $printerName
    if ($null -eq $printer) {
        return @{ Ok = $false; Text = "找不到打印机：$printerName" }
    }
    if ($printer.WorkOffline) {
        return @{ Ok = $false; Text = "$printerName：脱机" }
    }
    if ($printer.PrinterState -band 1) {
        return @{ Ok = $false; Text = "$printerName：暂停" }
    }
    if ($printer.PrinterState -band 2) {
        return @{ Ok = $false; Text = "$printerName：错误" }
    }
    if ($printer.PrinterState -band 16) {
        return @{ Ok = $false; Text = "$printerName：缺纸" }
    }
    if ($printer.PrinterState -band 128) {
        return @{ Ok = $false; Text = "$printerName：离线" }
    }
    if ($printer.PrinterState -band 1024) {
        return @{ Ok = $false; Text = "$printerName：需要处理" }
    }
    return @{ Ok = $true; Text = "$printerName：就绪" }
}

function Set-StatusLabel($label, $result) {
    if ($null -eq $label -or $null -eq $result) { return }
    $label.Text = $result.Text
    if ($result.Ok) {
        $label.ForeColor = [System.Drawing.Color]::FromArgb(32, 128, 64)
    } else {
        $label.ForeColor = [System.Drawing.Color]::FromArgb(200, 30, 30)
    }
}

function Write-Log($text) {
    $time = Get-Date -Format "HH:mm:ss"
    $logBox.AppendText("[$time] $text`r`n")
    $logBox.SelectionStart = $logBox.Text.Length
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-SalesSheetPrintSetup($sheet, $paper) {
    $xlLandscape = 2
    if ($paper -eq "third") {
        $printArea = "A1:H17"
        $left = 0.12
        $right = 0.12
        $top = 0.10
        $bottom = 0.10
    } else {
        $printArea = "A1:H25"
        $left = 0.18
        $right = 0.18
        $top = 0.14
        $bottom = 0.14
    }
    $sheet.PageSetup.PrintArea = $printArea
    $sheet.PageSetup.Orientation = $xlLandscape
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = 1
    $sheet.PageSetup.LeftMargin = $sheet.Application.InchesToPoints($left)
    $sheet.PageSetup.RightMargin = $sheet.Application.InchesToPoints($right)
    $sheet.PageSetup.TopMargin = $sheet.Application.InchesToPoints($top)
    $sheet.PageSetup.BottomMargin = $sheet.Application.InchesToPoints($bottom)
    $sheet.PageSetup.HeaderMargin = 0
    $sheet.PageSetup.FooterMargin = 0
    $sheet.PageSetup.CenterHorizontally = $true
    $sheet.PageSetup.CenterVertically = $false
}

function Set-PurchaseSheetPrintSetup($sheet) {
    $xlPortrait = 1
    $sheet.PageSetup.Orientation = $xlPortrait
    $sheet.PageSetup.Zoom = $false
    $sheet.PageSetup.FitToPagesWide = 1
    $sheet.PageSetup.FitToPagesTall = $false
    $sheet.PageSetup.LeftMargin = $sheet.Application.InchesToPoints(0.25)
    $sheet.PageSetup.RightMargin = $sheet.Application.InchesToPoints(0.25)
    $sheet.PageSetup.TopMargin = $sheet.Application.InchesToPoints(0.30)
    $sheet.PageSetup.BottomMargin = $sheet.Application.InchesToPoints(0.30)
    $sheet.PageSetup.HeaderMargin = 0
    $sheet.PageSetup.FooterMargin = 0
    $sheet.PageSetup.CenterHorizontally = $true
}

function Print-Workbook($filePath, $printerName, $jobType, $paper, $copies) {
    $excel = $null
    $workbook = $null
    $oldDefaultPrinter = $null
    try {
        $oldDefaultPrinter = Get-CimInstance Win32_Printer | Where-Object { $_.Default } | Select-Object -First 1
        $network = New-Object -ComObject WScript.Network
        $network.SetDefaultPrinter($printerName)

        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Open($filePath)
        foreach ($sheet in @($workbook.Worksheets)) {
            if ($jobType -eq "sales") {
                Set-SalesSheetPrintSetup $sheet $paper
            } else {
                Set-PurchaseSheetPrintSetup $sheet
            }
            Write-Log "打印：$([System.IO.Path]::GetFileName($filePath)) / $($sheet.Name) -> $printerName"
            $sheet.PrintOut($null, $null, $copies, $false)
        }
    } finally {
        if ($workbook -ne $null) {
            $workbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel -ne $null) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        if ($oldDefaultPrinter -and $oldDefaultPrinter.Name -and $oldDefaultPrinter.Name -ne $printerName) {
            try {
                $network = New-Object -ComObject WScript.Network
                $network.SetDefaultPrinter($oldDefaultPrinter.Name)
            } catch {}
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Test-ServerConnection {
    try {
        $cfg = Load-Config
        $url = "$($cfg.serverUrl.TrimEnd('/'))/api/print-agent/ping?token=$($cfg.token)"
        $res = Invoke-RestMethod -Uri $url -Method Get -TimeoutSec 8
        if ($res.ok) {
            return @{ Ok = $true; Text = "服务器连接：正常" }
        }
        return @{ Ok = $false; Text = "服务器连接：异常" }
    } catch {
        return @{ Ok = $false; Text = "服务器连接：失败" }
    }
}

function Refresh-Status {
    try {
        $cfg = Load-Config
        $salesPrinter = Resolve-Printer $cfg.salesPrinter "sales"
        $purchasePrinter = Resolve-Printer $cfg.purchasePrinter "purchase"
        Set-StatusLabel $serverStatusLabel (Test-ServerConnection)
        Set-StatusLabel $salesStatusLabel (Get-Printer-StatusText $salesPrinter)
        Set-StatusLabel $purchaseStatusLabel (Get-Printer-StatusText $purchasePrinter)
        Write-Log "检测完成"
    } catch {
        Write-Log "检测失败：$($_.Exception.Message)"
    }
}

function Create-TestWorkbook($jobType) {
    $fileName = if ($jobType -eq "sales") { "测试销售单.xlsx" } else { "测试采购单.xlsx" }
    $path = Join-Path $DownloadDir $fileName
    $excel = $null
    $workbook = $null
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $workbook = $excel.Workbooks.Add()
        $sheet = $workbook.Worksheets.Item(1)
        if ($jobType -eq "sales") {
            $sheet.Name = "销售单测试"
            $sheet.Range("A1:H1").Merge() | Out-Null
            $sheet.Range("A1").Value2 = "销售单测试页"
            $sheet.Range("A1").Font.Size = 18
            $sheet.Range("A1").Font.Bold = $true
            $sheet.Range("A3").Value2 = "如果 Epson 610K 打出这一页，说明销售单打印通道正常。"
            $sheet.Range("A5").Value2 = "纸张：二等分 241mm x 140mm"
            $sheet.Range("A7:H12").Borders.LineStyle = 1
            $sheet.Range("A7").Value2 = "序号"
            $sheet.Range("B7").Value2 = "品名"
            $sheet.Range("D7").Value2 = "数量"
            $sheet.Range("E7").Value2 = "单位"
            $sheet.Range("F7").Value2 = "单价"
            $sheet.Range("G7").Value2 = "金额"
            $sheet.Range("A8").Value2 = 1
            $sheet.Range("B8").Value2 = "测试商品"
            $sheet.Range("D8").Value2 = 1
            $sheet.Range("E8").Value2 = "斤"
            $sheet.Range("F8").Value2 = 1
            $sheet.Range("G8").Formula = "=D8*F8"
            $sheet.Columns.AutoFit() | Out-Null
            Set-SalesSheetPrintSetup $sheet "half"
        } else {
            $sheet.Name = "采购单测试"
            $sheet.Range("A1:D1").Merge() | Out-Null
            $sheet.Range("A1").Value2 = "采购单测试页"
            $sheet.Range("A1").Font.Size = 18
            $sheet.Range("A1").Font.Bold = $true
            $sheet.Range("A3").Value2 = "如果 HP 激光打印机打出这一页，说明采购单打印通道正常。"
            $sheet.Range("A5:D8").Borders.LineStyle = 1
            $sheet.Range("A5").Value2 = "食材分类"
            $sheet.Range("B5").Value2 = "食材名称"
            $sheet.Range("C5").Value2 = "采购数量"
            $sheet.Range("D5").Value2 = "计量单位"
            $sheet.Range("A6").Value2 = "蔬菜类"
            $sheet.Range("B6").Value2 = "测试商品"
            $sheet.Range("C6").Value2 = 1
            $sheet.Range("D6").Value2 = "斤"
            $sheet.Columns.AutoFit() | Out-Null
            Set-PurchaseSheetPrintSetup $sheet
        }
        if (Test-Path $path) { Remove-Item $path -Force }
        $workbook.SaveAs($path)
        return $path
    } finally {
        if ($workbook -ne $null) {
            $workbook.Close($false)
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        }
        if ($excel -ne $null) {
            $excel.Quit()
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        }
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
    }
}

function Print-TestPage($jobType) {
    try {
        $cfg = Load-Config
        $printer = if ($jobType -eq "sales") {
            Resolve-Printer $cfg.salesPrinter "sales"
        } else {
            Resolve-Printer $cfg.purchasePrinter "purchase"
        }
        $status = Get-Printer-StatusText $printer
        if (!$status.Ok) {
            $targetLabel = if ($jobType -eq "sales") { $salesStatusLabel } else { $purchaseStatusLabel }
            Set-StatusLabel $targetLabel $status
            throw $status.Text
        }
        $path = Create-TestWorkbook $jobType
        $paper = if ($jobType -eq "sales") { "half" } else { "a4" }
        Print-Workbook $path $printer $jobType $paper 1
        Write-Log "测试打印已发送：$printer"
    } catch {
        Write-Log "测试打印失败：$($_.Exception.Message)"
    }
}

function Complete-Job($cfg, $jobId, $ok, $errorText) {
    $uri = "$($cfg.serverUrl.TrimEnd('/'))/api/print-agent/jobs/$jobId/complete?token=$($cfg.token)"
    $body = @{ ok = $ok; error = $errorText } | ConvertTo-Json -Compress
    Invoke-RestMethod -Uri $uri -Method Post -ContentType "application/json; charset=utf-8" -Body $body | Out-Null
}

function Download-JobFile($cfg, $job) {
    $safeTitle = [Regex]::Replace(($job.title | Out-String).Trim(), '[\\/:*?"<>|]', '_')
    if (!$safeTitle) { $safeTitle = $job.id }
    $path = Join-Path $DownloadDir "$safeTitle.xlsx"
    $url = "$($cfg.serverUrl.TrimEnd('/'))$($job.downloadUrl)"
    Invoke-WebRequest -Uri $url -OutFile $path
    return $path
}

function Handle-Jobs($jobs) {
    if ($null -eq $jobs) { return }
    $cfg = Load-Config
    foreach ($job in @($jobs)) {
        try {
            $printer = if ($job.type -eq "sales") {
                Resolve-Printer $cfg.salesPrinter "sales"
            } else {
                Resolve-Printer $cfg.purchasePrinter "purchase"
            }
            if (!$printer) {
                throw "没有找到可用打印机"
            }
            Write-Log "收到任务：$($job.title)"
            $filePath = Download-JobFile $cfg $job
            Print-Workbook $filePath $printer $job.type $job.paper ([int]$cfg.copies)
            Complete-Job $cfg $job.id $true ""
            Write-Log "完成：$($job.title)"
        } catch {
            Write-Log "失败：$($job.title) / $($_.Exception.Message)"
            try { Complete-Job $cfg $job.id $false $_.Exception.Message } catch {}
        }
    }
}

function Start-EventLoop {
    while ($true) {
        $cfg = Load-Config
        $url = "$($cfg.serverUrl.TrimEnd('/'))/api/print-agent/events?token=$($cfg.token)"
        Write-Log "连接服务器：$($cfg.serverUrl)"
        try {
            Set-StatusLabel $serverStatusLabel @{ Ok = $true; Text = "服务器连接：正常，正在接收任务" }
            $request = [System.Net.HttpWebRequest]::Create($url)
            $request.Method = "GET"
            $request.Accept = "text/event-stream"
            $request.Timeout = 300000
            $response = $request.GetResponse()
            $stream = $response.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream, [System.Text.Encoding]::UTF8)
            while (!$reader.EndOfStream) {
                $line = $reader.ReadLine()
                if ($line -and $line.StartsWith("data: ")) {
                    $json = $line.Substring(6)
                    if ($json -and $json -ne '{"message":"connected"}') {
                        $jobs = $json | ConvertFrom-Json
                        Handle-Jobs $jobs
                    }
                }
                [System.Windows.Forms.Application]::DoEvents()
            }
        } catch {
            Set-StatusLabel $serverStatusLabel @{ Ok = $false; Text = "服务器连接：断开，正在重连" }
            Write-Log "连接断开，5秒后重连：$($_.Exception.Message)"
            Start-Sleep -Seconds 5
        }
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "酒店打印助手"
$form.Size = New-Object System.Drawing.Size(720, 620)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$statusLabel = New-Object System.Windows.Forms.Label
$statusLabel.Text = "启动后会自动接收服务器发送的销售单、采购单，并按配置切换打印机。"
$statusLabel.Location = New-Object System.Drawing.Point(18, 18)
$statusLabel.Size = New-Object System.Drawing.Size(660, 28)
$form.Controls.Add($statusLabel)

$serverStatusLabel = New-Object System.Windows.Forms.Label
$serverStatusLabel.Text = "服务器连接：待检测"
$serverStatusLabel.Location = New-Object System.Drawing.Point(18, 54)
$serverStatusLabel.Size = New-Object System.Drawing.Size(660, 24)
$form.Controls.Add($serverStatusLabel)

$salesStatusLabel = New-Object System.Windows.Forms.Label
$salesStatusLabel.Text = "销售单打印机：待检测"
$salesStatusLabel.Location = New-Object System.Drawing.Point(18, 82)
$salesStatusLabel.Size = New-Object System.Drawing.Size(660, 24)
$form.Controls.Add($salesStatusLabel)

$purchaseStatusLabel = New-Object System.Windows.Forms.Label
$purchaseStatusLabel.Text = "采购单打印机：待检测"
$purchaseStatusLabel.Location = New-Object System.Drawing.Point(18, 110)
$purchaseStatusLabel.Size = New-Object System.Drawing.Size(660, 24)
$form.Controls.Add($purchaseStatusLabel)

$openConfigBtn = New-Object System.Windows.Forms.Button
$openConfigBtn.Text = "打开配置"
$openConfigBtn.Location = New-Object System.Drawing.Point(18, 145)
$openConfigBtn.Size = New-Object System.Drawing.Size(120, 36)
$form.Controls.Add($openConfigBtn)

$openFolderBtn = New-Object System.Windows.Forms.Button
$openFolderBtn.Text = "打开接收目录"
$openFolderBtn.Location = New-Object System.Drawing.Point(154, 145)
$openFolderBtn.Size = New-Object System.Drawing.Size(140, 36)
$form.Controls.Add($openFolderBtn)

$checkBtn = New-Object System.Windows.Forms.Button
$checkBtn.Text = "立即检测"
$checkBtn.Location = New-Object System.Drawing.Point(310, 145)
$checkBtn.Size = New-Object System.Drawing.Size(110, 36)
$form.Controls.Add($checkBtn)

$testSalesBtn = New-Object System.Windows.Forms.Button
$testSalesBtn.Text = "测试销售单"
$testSalesBtn.Location = New-Object System.Drawing.Point(436, 145)
$testSalesBtn.Size = New-Object System.Drawing.Size(115, 36)
$form.Controls.Add($testSalesBtn)

$testPurchaseBtn = New-Object System.Windows.Forms.Button
$testPurchaseBtn.Text = "测试采购单"
$testPurchaseBtn.Location = New-Object System.Drawing.Point(567, 145)
$testPurchaseBtn.Size = New-Object System.Drawing.Size(115, 36)
$form.Controls.Add($testPurchaseBtn)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(18, 198)
$logBox.Size = New-Object System.Drawing.Size(665, 340)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$form.Controls.Add($logBox)

$openConfigBtn.Add_Click({ Start-Process notepad.exe $ConfigPath })
$openFolderBtn.Add_Click({ Start-Process explorer.exe $DownloadDir })
$checkBtn.Add_Click({ Refresh-Status })
$testSalesBtn.Add_Click({ Print-TestPage "sales" })
$testPurchaseBtn.Add_Click({ Print-TestPage "purchase" })

$form.Add_Shown({
    $cfg = Load-Config
    $salesPrinter = Resolve-Printer $cfg.salesPrinter "sales"
    $purchasePrinter = Resolve-Printer $cfg.purchasePrinter "purchase"
    Write-Log "销售单打印机：$salesPrinter"
    Write-Log "采购单打印机：$purchasePrinter"
    Refresh-Status
    Start-EventLoop
})

[System.Windows.Forms.Application]::Run($form)
