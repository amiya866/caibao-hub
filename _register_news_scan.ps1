$py = "C:\Users\Yitian Shen\AppData\Local\Programs\Python\Python314\pythonw.exe"
$act = New-ScheduledTaskAction -Execute $py -Argument '"D:\Kimi\caibao-hub\_caibao_news_scan.py"' -WorkingDirectory "D:\Kimi\caibao-hub"
$trigs = @()
foreach ($hm in @("09:35","13:35","17:35","21:35")) {
  $trigs += New-ScheduledTaskTrigger -Daily -At $hm
}
$set = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Minutes 30)
Register-ScheduledTask -TaskName "CaibaoNewsScan" -Action $act -Trigger $trigs -Settings $set -Description "caibao-hub news.json auto scan from workbench snapshot, build+push" -Force
Get-ScheduledTask -TaskName "CaibaoNewsScan" | Select-Object TaskName, State
