Add-Type -AssemblyName System.Windows.Forms
$label = New-Object System.Windows.Forms.Label
$label.AutoSize = $true
$form = New-Object System.Windows.Forms.Form
$form.Controls.Add($label)
$form.TopMost = $true

#remove existing timers
foreach ($t in $timer) {
    $t.Stop()
    $t.Dispose()
}
$timer = @()
#remove existing timers end

$timer = New-Object System.Windows.Forms.Timer
$timer.Interval = 5000 # close windows every 5 seconds
$timer.Add_Tick({
    $form.Close()
})
$timer.Start()

while ($true) {
    $spooler = Get-WmiObject -Class Win32_PrintJob
    if ($spooler -eq $null) {
        $wshell = New-Object -ComObject Wscript.Shell
        $result = $wshell.Popup("Your computer will shutdown after 5 seconds.", 5, "shutdown", 0x1)
        if ($result -eq 2) {
            break
        }
        #Start-Sleep -Seconds 5
        Stop-Computer
    } else {
        $printer = Get-WmiObject -Class Win32_Printer | Where-Object {$_.Default -eq $true} #get printer status
        $printerStatus = "NULL"
        switch($printer.PrinterStatus){
            1 { $printerStatus = "[1] Other"}
            2 { $printerStatus = "[2] Unknown"}
            3 { $printerStatus = "[3] Idle"}
            4 { $printerStatus = "[4] Printing"}
            5 { $printerStatus = "[5] Warm up"}
            6 { $printerStatus = "[6] Stopped Printing"}
            7 { $printerStatus = "[7] Offline"}
            default { $printerStatus = "$($printer.PrinterStatus): Cannot find response code at https://learn.microsoft.com/en-us/windows/win32/cimwin32prov/win32-printer."}
        }
        $text = "Default printer: $($printer.Name)`nStatus: $printerStatus`n`n"
        
        $text += "Printing File name:`n"
        foreach ($job in $spooler) {
            if ($job.JobStatus -eq "Printing") {
                $text += "`t$($job.Document)`n"
            }
        }
        if ([int]$spooler.Count -eq 0) {
            $text += "Number of tasks remaining: 0"
        } else {
            $text += "Number of tasks remaining: $($spooler.Count -1)"
        }
        
        $label.Text = $text
        if (!$form.Visible) {
            $form.ShowDialog()
        }
    }
    Start-Sleep -Seconds 1
}
