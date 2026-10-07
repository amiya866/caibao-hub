$log = 'D:\Kimi\caibao-hub\_watch_task.log'
"=== $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') ===" | Out-File $log -Append -Encoding utf8
python 'D:\Kimi\caibao-hub\_caibao_watch.py' >> $log 2>&1
"rc=$LASTEXITCODE" | Out-File $log -Append -Encoding utf8
