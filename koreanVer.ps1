Add-Type -AssemblyName System.Windows.Forms
$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
$form = New-Object System.Windows.Forms.Form
$form.Controls.Add($label)
$form.TopMost = $true
$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000 # 5초마다 창을 닫음
$timer.Add_Tick({
    $form.Close()
})
$timer.Start()

while ($true) {
    $spooler = Get-WmiObject -Class Win32_PrintJob
    if ($spooler -eq $null) {
        $wshell = New-Object -ComObject Wscript.Shell
        $result = $wshell.Popup("컴퓨터가 5초 후에 종료됩니다.", 5, "종료", 0x1)
        if ($result -eq 2) {
            break
        }
        Start-Sleep -Seconds 5
        Stop-Computer
    } else {
        $printer = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Default -eq $true}
        $printerStatus = "NULL"
        switch($printer.PrinterStatus){
            1 { $printerStatus = "기타"}
            2 { $printerStatus = "알 수 없음"}
            3 { $printerStatus = "유휴 상태"}
            4 { $printerStatus = "인쇄"}
            5 { $printerStatus = "준비"}
            6 { $printerStatus = "인쇄 중지됨"}
            7 { $printerStatus = "오프라인"}
            default { $printerStatus = "알 수 없음"}
        $text = "기본 프린터: $($printer.Name)`n상태: $printerStatus`n`n"
        
        $text += "인쇄 중인 파일:`n"
        foreach ($job in $spooler) {
            if ($job.JobStatus -eq "Printing") {
                $text += "`t$($job.Document)`n"
            }
        }
        if ([int]$spooler.Count -eq 0) {
            $text += "남은 작업 개수: 0"
        } else {
            $text += "남은 작업 개수: $($spooler.Count -1)"
        }
        
        $label.Text = $text
        if (!$form.Visible) {
            $form.ShowDialog()
        }
    }
    Start-Sleep -Seconds 1
}