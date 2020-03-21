$sb = { 
    Import-Module –Name "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync" -Verbose
    Start-ADSyncSyncCycle -PolicyType Delta
}

$AADserv = 'ECQCON01'

Invoke-Command -ComputerName $AADserv -ScriptBlock $sb