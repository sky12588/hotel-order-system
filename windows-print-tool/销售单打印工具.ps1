Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = "Stop"

function Get-InstalledPrinters {
    try {
        return @(Get-CimInstance Win32_Printer | Sort-Object Name | ForEach-Object { $_.Name })
    } catch {
        return @()
    }
}

function Write-Log($text) {
    $time = Get-Date -Format "HH:mm:ss"
    $logBox.AppendText("[$time] $text`r`n")
    $logBox.SelectionStart = $logBox.Text.Length
    $logBox.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-SalesSheetPrintSetup($sheet, $mode) {
    $xlLandscape = 2
    $xlPaperUser = 256

    if ($mode -eq "third") {
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
    try { $sheet.PageSetup.PaperSize = $xlPaperUser } catch {}
}

function Print-SalesWorkbook($filePath, $printerName, $mode, $copies) {
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
            Set-SalesSheetPrintSetup $sheet $mode
            Write-Log "打印：$([System.IO.Path]::GetFileName($filePath)) / $($sheet.Name)"
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

$form = New-Object System.Windows.Forms.Form
$form.Text = "销售单快速打印工具"
$form.Size = New-Object System.Drawing.Size(620, 520)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Microsoft YaHei UI", 10)

$fileLabel = New-Object System.Windows.Forms.Label
$fileLabel.Text = "销售单 Excel 文件"
$fileLabel.Location = New-Object System.Drawing.Point(18, 22)
$fileLabel.Size = New-Object System.Drawing.Size(160, 28)
$form.Controls.Add($fileLabel)

$fileBox = New-Object System.Windows.Forms.TextBox
$fileBox.Location = New-Object System.Drawing.Point(18, 52)
$fileBox.Size = New-Object System.Drawing.Size(460, 30)
$fileBox.ReadOnly = $true
$form.Controls.Add($fileBox)

$browseBtn = New-Object System.Windows.Forms.Button
$browseBtn.Text = "选择文件"
$browseBtn.Location = New-Object System.Drawing.Point(490, 50)
$browseBtn.Size = New-Object System.Drawing.Size(95, 34)
$form.Controls.Add($browseBtn)

$printerLabel = New-Object System.Windows.Forms.Label
$printerLabel.Text = "打印机"
$printerLabel.Location = New-Object System.Drawing.Point(18, 100)
$printerLabel.Size = New-Object System.Drawing.Size(160, 28)
$form.Controls.Add($printerLabel)

$printerBox = New-Object System.Windows.Forms.ComboBox
$printerBox.Location = New-Object System.Drawing.Point(18, 130)
$printerBox.Size = New-Object System.Drawing.Size(360, 30)
$printerBox.DropDownStyle = "DropDownList"
$printers = Get-InstalledPrinters
foreach ($printer in $printers) { [void]$printerBox.Items.Add($printer) }
$defaultPrinter = Get-CimInstance Win32_Printer | Where-Object { $_.Default } | Select-Object -First 1
if ($defaultPrinter) {
    $printerBox.SelectedItem = $defaultPrinter.Name
} elseif ($printerBox.Items.Count -gt 0) {
    $printerBox.SelectedIndex = 0
}
$form.Controls.Add($printerBox)

$modeLabel = New-Object System.Windows.Forms.Label
$modeLabel.Text = "纸张模式"
$modeLabel.Location = New-Object System.Drawing.Point(400, 100)
$modeLabel.Size = New-Object System.Drawing.Size(160, 28)
$form.Controls.Add($modeLabel)

$modeBox = New-Object System.Windows.Forms.ComboBox
$modeBox.Location = New-Object System.Drawing.Point(400, 130)
$modeBox.Size = New-Object System.Drawing.Size(185, 30)
$modeBox.DropDownStyle = "DropDownList"
[void]$modeBox.Items.Add("二等分连续纸 241x140")
[void]$modeBox.Items.Add("三等分连续纸")
$modeBox.SelectedIndex = 0
$form.Controls.Add($modeBox)

$copiesLabel = New-Object System.Windows.Forms.Label
$copiesLabel.Text = "份数"
$copiesLabel.Location = New-Object System.Drawing.Point(18, 178)
$copiesLabel.Size = New-Object System.Drawing.Size(80, 28)
$form.Controls.Add($copiesLabel)

$copiesBox = New-Object System.Windows.Forms.NumericUpDown
$copiesBox.Location = New-Object System.Drawing.Point(70, 176)
$copiesBox.Size = New-Object System.Drawing.Size(90, 30)
$copiesBox.Minimum = 1
$copiesBox.Maximum = 5
$copiesBox.Value = 1
$form.Controls.Add($copiesBox)

$hint = New-Object System.Windows.Forms.Label
$hint.Text = "提示：第一次请在 Windows 打印机首选项里把 Epson 610K 的二等分/三等分纸张保存好。以后这里选择文件和打印机即可。"
$hint.Location = New-Object System.Drawing.Point(18, 218)
$hint.Size = New-Object System.Drawing.Size(565, 52)
$hint.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$form.Controls.Add($hint)

$printBtn = New-Object System.Windows.Forms.Button
$printBtn.Text = "开始打印"
$printBtn.Location = New-Object System.Drawing.Point(18, 282)
$printBtn.Size = New-Object System.Drawing.Size(140, 42)
$printBtn.BackColor = [System.Drawing.Color]::FromArgb(22, 119, 255)
$printBtn.ForeColor = [System.Drawing.Color]::White
$form.Controls.Add($printBtn)

$openPrinterBtn = New-Object System.Windows.Forms.Button
$openPrinterBtn.Text = "打开打印机设置"
$openPrinterBtn.Location = New-Object System.Drawing.Point(178, 282)
$openPrinterBtn.Size = New-Object System.Drawing.Size(160, 42)
$form.Controls.Add($openPrinterBtn)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(18, 342)
$logBox.Size = New-Object System.Drawing.Size(567, 120)
$logBox.Multiline = $true
$logBox.ScrollBars = "Vertical"
$logBox.ReadOnly = $true
$form.Controls.Add($logBox)

$browseBtn.Add_Click({
    $dialog = New-Object System.Windows.Forms.OpenFileDialog
    $dialog.Title = "选择销售单 Excel"
    $dialog.Filter = "Excel 文件 (*.xlsx)|*.xlsx"
    $dialog.Multiselect = $false
    if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $fileBox.Text = $dialog.FileName
    }
})

$openPrinterBtn.Add_Click({
    Start-Process "control.exe" "printers"
})

$printBtn.Add_Click({
    try {
        if (-not (Test-Path $fileBox.Text)) {
            [System.Windows.Forms.MessageBox]::Show("请先选择销售单 Excel 文件。", "提示") | Out-Null
            return
        }
        if (-not $printerBox.SelectedItem) {
            [System.Windows.Forms.MessageBox]::Show("请先选择打印机。", "提示") | Out-Null
            return
        }
        $mode = if ($modeBox.SelectedIndex -eq 1) { "third" } else { "half" }
        $printBtn.Enabled = $false
        Write-Log "开始打印，打印机：$($printerBox.SelectedItem)"
        Print-SalesWorkbook $fileBox.Text $printerBox.SelectedItem $mode ([int]$copiesBox.Value)
        Write-Log "打印任务已发送。"
        [System.Windows.Forms.MessageBox]::Show("打印任务已发送。", "完成") | Out-Null
    } catch {
        Write-Log "失败：$($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("打印失败：$($_.Exception.Message)", "错误") | Out-Null
    } finally {
        $printBtn.Enabled = $true
    }
})

Write-Log "工具已启动。"
[void]$form.ShowDialog()
