Add-Type -AssemblyName System.Windows.Forms


$delayBeforeStart = 5       
$times = 20                 
$interval = 10             


$messages = @(
    @{Text=":thumbsdown:"; Enters=2},
    @{Text="<3"; Enters=1}
    #@{Text="text"; Enters=1}
    #@{Text="text"; Enters=1}
    #@{Text="text"; Enters=1}
)





$global:stopFlag = $false
$form = New-Object System.Windows.Forms.Form
$form.WindowState = "Minimized"
$form.ShowInTaskbar = $false

$form.Add_KeyDown({
    if ($_.Control -and $_.KeyCode -eq "Q") {
        $global:stopFlag = $true
    }
})

$null = $form.CreateControl()

Write-Host "Switch to your Telegram window. Starting in $delayBeforeStart seconds..."
Start-Sleep -Seconds $delayBeforeStart

$wshell = New-Object -ComObject WScript.Shell


for ($i = 1; $i -le $times; $i++) {
    if ($global:stopFlag) {
        Write-Host "Stopped by user (Ctrl+Q)"
        break
    }

    $msgObj = Get-Random -InputObject $messages

    
    $wshell.SendKeys($msgObj.Text)

    for ($e=1; $e -le $msgObj.Enters; $e++) {
        $wshell.SendKeys("{ENTER}")
        Start-Sleep -Milliseconds 50
    }

    Start-Sleep -Milliseconds $interval
}

Write-Host "Done."
