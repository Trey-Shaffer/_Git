Param( [string]$GitPath = "D:\Program Files\Git\bin\git.exe" )
CLS
Get-Date

# Alias Git
New-Alias -Name git -Value $GitPath

"Done..."
