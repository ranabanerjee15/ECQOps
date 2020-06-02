# below command will install RSAT (run ise as an Administrator)
Install-WindowsFeature -IncludeAllSubFeature RSAT

#Install Nuget Package
Install-Module PowershellGet -Force -Confirm:$false -AllowClobber -SkipPublisherCheck

# Installs Other required Modules
'ImportExcel', 'EnhancedHTML2', 'AzureAD', 'MsOnline', 'Az' |
ForEach-Object { Install-Module $_ -AllowClobber -Force -Confirm:$false -SkipPublisherCheck -Verbose }

#Install Chocolatey Package Manager
Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('https://chocolatey.org/install.ps1'))

# Install Windows Terminal
choco install microsoft-windows-terminal

#Install PowerShell

iex "& { $(irm https://aka.ms/install-powershell.ps1) } -UseMSI"

Invoke-Expression "& { $(Invoke-Restmethod https://aka.ms/install-powershell.ps1) } -UseMSI -Preview"

# Secrets Management Module
Install-module 'Microsoft.PowerShell.SecretManagement' -SkipPublisherCheck -AllowClobber -Confirm:$false -Force -AllowPrerelease -Verbose

# Compatiblity module
Install-Module WindowsCompatibility 