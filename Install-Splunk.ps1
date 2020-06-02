# Get the ID and security principal of the current user account
$myWindowsID = [System.Security.Principal.WindowsIdentity]::GetCurrent();
$myWindowsPrincipal = New-Object System.Security.Principal.WindowsPrincipal($myWindowsID);

# Get the security principal for the administrator role
$adminRole = [System.Security.Principal.WindowsBuiltInRole]::Administrator;

# Check to see if we are currently running as an administrator
if ($myWindowsPrincipal.IsInRole($adminRole))
{
    # We are running as an administrator, so change the title and background colour to indicate this
    $Host.UI.RawUI.WindowTitle = $myInvocation.MyCommand.Definition + "(Elevated)";
    $Host.UI.RawUI.BackgroundColor = "black";
    Clear-Host;
}
else {
    # We are not running as an administrator, so relaunch as administrator

    # Create a new process object that starts PowerShell
    $newProcess = New-Object System.Diagnostics.ProcessStartInfo "PowerShell";

    # Specify the current script path and name as a parameter with added scope and support for scripts with spaces in it's path
    $newProcess.Arguments = "& '" + $script:MyInvocation.MyCommand.Path + "'"

    # Indicate that the process should be elevated
    $newProcess.Verb = "runas";

    # Start the new process
    [System.Diagnostics.Process]::Start($newProcess);

    # Exit from the current, unelevated, process
    Exit;
}

# Run your code that needs to be elevated here...
Write-Host "Press any key to continue...";
$null = $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown");

$args = @(
    "/i \\ecqfps01\netshare$\Common\zTempInstallers\splunkforwarder-7.3.1-bd63e13aa157-x64-release.msi"
    "AGREETOLICENSE=Yes"
    "DEPLOYMENT_SERVER=ecqsplunk.australiaeast.cloudapp.azure.com:8089"
    "/l*v C:\SplunkInstall.log"
    "/quiet"
)
Start-Process msiexec.exe -Verb RunAs -Wait -ArgumentList $args -Verbose

Write-Host "Installed Splunk Forwarder" -ForegroundColor Cyan

$splunkConfigFile = "$env:ProgramFiles" +"\SplunkUniversalForwarder\etc\system\local\server.conf"

$append = @"

[deployment]
pass4SymmKey = 53f9c90785aac5ecb731d44bda886686
"@


$append | Out-File -Append -Encoding utf8 -FilePath $splunkConfigFile

Write-Host "Appended Splunk Conf File" -ForegroundColor Cyan
Start-Sleep -Seconds 10
Restart-Service -Name SplunkForwarder -Verbose

Write-Host "Restarted splunk Forwarder Service" -ForegroundColor Cyan

notepad $splunkConfigFile
notepad C:\SplunkInstall.log
