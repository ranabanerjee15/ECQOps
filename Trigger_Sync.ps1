$sb = {
    Import-Module –Name "C:\Program Files\Microsoft Azure AD Sync\Bin\ADSync" -Verbose
    Start-ADSyncSyncCycle -PolicyType Delta -Verbose
}

Invoke-Command -ComputerName ECQCON01 -ScriptBlock $sb