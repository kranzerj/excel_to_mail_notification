#script does not work at older Windows Server like 2016
# use powershell 7 (C:\Program Files\PowerShell\7\pwsh.exe)
# install csvkit, pandas and chardet (pip install csvkit pandas chardet )




Set-Location -Path "C:\TS Benachrichtiger\System und Update Liste"

Remove-Item "System und Update Liste.xlsx"

Remove-Item "System und Update Liste.csv"

Remove-Item "Übersicht Handlungsbedarf.csv"

Remove-Item "System und Update Liste.temp.csv"





$DownloadUrl = "https://kundenurl.sharepoint.com/:x:/s/bliblablup/sharingURL7Q?download=1" 
$OutputFile = "System und Update Liste.xlsx" 

Invoke-WebRequest -Uri $DownloadUrl -OutFile $OutputFile 

Start-Sleep -Seconds 64


in2csv "System und Update Liste.xlsx" > "System und Update Liste.temp.csv"


Start-Sleep -Seconds 32

python .\convert_csv.py

Start-Sleep -Seconds 32


python logik.py 
