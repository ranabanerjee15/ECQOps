<#	
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.172
	 Created on:   	17/02/2020 2:56 PM
	 Created by:   	Rana Banerjee
	 Organization: 	ECQ
	 Filename:     	ECQOps.psm1
	-------------------------------------------------------------------------
	 Module Name: ECQOps
	===========================================================================
#>

function Connect-365
{
<#
	.SYNOPSIS
		Connect to Office 365 in PSSession
	
	.DESCRIPTION
		Connect to office 365 using App Password or your regular password. Note currently this function does not cater for MFA authentication. If you have enabled MFA please generate and use APP password.
	
	.PARAMETER credential
		This is the credential which you will use to connect to Exchange Online.
	
	.PARAMETER UseIEProxy
		if the connection needs to go via proxy settings, this switch will use the default proxy settings from the Internet Explorer
	
	.EXAMPLE
		PS C:\> Connect-QhO365 -Credential 'Adminuser@domain.onmicrosoft.com'
	
	.EXAMPLE
		PS C:\> Connect-QhO365 -Credential 'Adminuser@domain.onmicrosoft.com' -UseIEProxy
	
	.EXAMPLE
		PS C:\> Connect-QhO365 -Credential (Get-Credential) -UseIEProxy
	
	.EXAMPLE
		PS C:\> Connect-QhO365 -Credential (Get-Credential)
	
	.EXAMPLE
		PS C:\> Connect-QhO365 -Credential $cred
	
	.NOTES
		In the above example $cred can be a variable which has your credentials stored.
#>
	
	param
	(
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty,
		[Switch]$UseIEProxy
	)
	try
	{
		$FormatEnumerationLimit = -1
		#Write-Host "INFO : Trying to Connect to Office 365" -ForegroundColor Cyan
		
		if ($credential -eq $null)
		{
			$credential = Get-Credential -Message "Enter your Credentials" -ErrorAction Stop
		}
		$paramNewPSSession = @{
			ConfigurationName = 'Microsoft.Exchange'
			ConnectionUri	  = 'https://outlook.office365.com/powershell-liveid/'
			Credential	      = $credential
			Authentication    = 'Basic'
			AllowRedirection  = $true
		}
		if ($UseIEProxy)
		{
			$proxysettings = New-PSSessionOption -ProxyAccessType IEConfig
			$paramNewPSSession.Add('SessionOption', "$proxySettings")
		}
		
		#Write-Host "$($paramNewPSSession | out-string)" -ForegroundColor Cyan
		
		$ExoSession = New-PSSession @paramNewPSSession
		
		$paramImportModule = @{
			ModuleInfo = (Import-PSSession $ExoSession -AllowClobber -DisableNameChecking)
			Global	   = $true
			ErrorAction = 'Stop'
			WarningAction = 'SilentlyContinue'
			Prefix	   = 'EXO'
		}
		Import-Module @paramImportModule
		#Import-Module MsOnline -ErrorAction Stop -Global
		
		#Write-Host "SUCCESS : Successfully Connected to Office 365" -ForegroundColor Green
	}
	catch
	{
		Write-host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
	}
}

function Set-FunctionTemplate #Template
{
<#
	.SYNOPSIS
		A brief description of the Set-QHFunctionTemplate function.
	
	.DESCRIPTION
		A detailed description of the Set-QHFunctionTemplate function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Set-QHFunctionTemplate -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true)]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
			}
			catch
			{
				
			}
			finally
			{
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Doing Some Processing'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Doing Some Processing' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function get-LastADSync
{
<#
	.SYNOPSIS
		Gets Last AD sync.
	
	.DESCRIPTION
		Gets last AD sync time stamp.
	
	.EXAMPLE
		PS C:\> get-LastADSync
	
	.NOTES
		Additional information about the function.
#>
	param (
	)
	begin
	{
		
	}
	process
	{
		try
		{
			$MsolInfo = Get-MsolCompanyInformation -ErrorAction Stop
			If ($MsolInfo.DirectorySynchronizationEnabled)
			{
				$lastSync = $msolInfo.LastDirSyncTime.ToLocalTime()
				$now = (get-date -ErrorAction Stop).ToLocalTime()
				$Duration = $now - $LastSync
				$durationMinsRaw = $duration.TotalMinutes
				$durationMins = [math]::Round($durationMinsRaw)
				
				$obj = [PsCustomObject][Ordered]@{
					CurrentTime	    = $now
					LastAADSyncTime = $lastSync
					TimeElapsed	    = "$durationMins Mins Ago"
					NextScheduledAADSync = $lastSync.AddMinutes(30)
					TimeToNextAADSync = "$([math]::Round(($lastSync.AddMinutes(30) - $now).TotalMinutes)) Mins Remaining"
					
				}
				Write-Output $obj
			}
			else
			{
				throw "ADSync is NOT ENABLED. [ ADSync Enabled System Status : $($MsolInfo.DirectorySynchronizationEnabled) ]"
			}
			
			
		}
		catch
		{
			#$ErrorMsg = "ERROR : $($MyInvocation.InvocationName) `t`t$($error[0].Exception.Message)"
			Write-Host "ERROR : $($_.Exception.Message)" -ForegroundColor Magenta
		}
	}
	end
	{
		
	}
}

function Get-365LicenseStatus
{
<#
	.SYNOPSIS
		Gets Licensing Information for an Office 365 User.
	
	.DESCRIPTION
		This function gets detailed licensing information for any given user or Users. It uses the MSOnline module to get the information.
	
	.PARAMETER UserPrincipalName
		This is the UserPrincipalName of any user or users in Office 365.
	
	.PARAMETER ShowProgress
		This Switch shows progress if there are set or array of users.
	
	.EXAMPLE
		PS C:\> Get-365LicenseStatus -UserPrincipalName 'user1@contoso.com'
	
	.EXAMPLE
		PS C:\> Get-365LicenseStatus -UserPrincipalName 'user1@contoso.com','user1@contoso.com'
	
	.EXAMPLE
		PS C:\> Get-365LicenseStatus -UserPrincipalName (Import-csv users.csv) -ShowProgress
	
	.EXAMPLE
		PS C:\> Get-MsolUser -All | Get-365LicenseStatus
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$Plans = Get-MsolAccountSku -ErrorAction Stop | Select-Object -ExpandProperty AccountSkuId
		$Company = $Plans[0].Split(':')[0]
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				
				$prop = [ordered]@{
					UserPrincipalName = $UPN
				}
				
				foreach ($Plan in $Plans)
				{
					$lic = $msol.Licenses.Where({ $Plan -eq ($_ | Select-Object AccountSkuId | Select-Object -ExpandProperty AccountSkuId) })
					if ($lic -ne $null)
					{
						$licobj = $lic | Select-Object @{
							n = 'AccountSku'; e = {
								switch ($lic | Select-Object AccountSkuid | Select-Object -ExpandProperty AccountSkuid)
								{
									"$($Company)):ENTERPRISEPACK" { 'E3' }
									"$($Company):STANDARDPACK" { 'E1' }
									"$($Company):MFA_STANDALONE" { 'MFA' }
									"$($Company):POWER_BI_STANDARD" { 'PowerBi Free' }
									"$($Company):EXCHANGEENTERPRISE" { 'ExchangeOnlinePlan2' }
									"$($Company):EMS" { 'Enterprise Mobility Security' }
									"$($Company):FLOW_FREE" { 'Microsoft Flow Free' }
									"$($Company):POWERAPPS_INDIVIDUAL_USER" { 'PowerApps and Logic Flows' }
									"$($Company):MCOEV" { 'Phone System' }
									"$($Company):POWER_BI_PRO" { 'PowerBi Pro' }
									"$($Company):POWER_BI_ADDON" { 'Power BI for Office 365 Add-On' }
									"$($Company):POWER_BI_INDIVIDUAL_USER" { 'Power BI Individual User' }
									"$($Company):ENTERPRISEWITHSCAL" { 'Enterprise Plan E4' }
									"$($Company):PROJECTONLINE_PLAN_1" { 'Project Online' }
									"$($Company):PROJECTCLIENT" { 'Project Pro for Office 365' }
									"$($Company):VISIOCLIENT" { 'Visio Pro Online' }
									"$($Company):STREAM" { 'Microsoft Stream' }
									"$($Company):POWERAPPS_VIRAL" { 'Microsoft Power Apps & Flow' }
									"$($Company):PROJECTESSENTIALS" { 'Project Lite' }
									"$($Company):PROJECTPROFESSIONAL" { 'Project Professional' }
									"$($Company):SPZA_IW" { 'App Connect' }
									"$($Company):PBI_PREMIUM_P1_ADDON" { 'Power Bi Premium' }
									"$($Company):DYN365_ENTERPRISE_P1_IW" { 'Dynamics 365 P1 Trial for Information Workers' }
									"$($Company):WINDOWS_STORE" { 'Windows Store for Business' }
									"$($Company):DEVELOPERPACK" { 'Developer Pack' }
									default { "$_" }
								}
							}
						},
													   @{
							n								  = 'Assignment'; e = {
								$lic.GroupsAssigningLicense.Guid | ForEach-Object {
									if ($_ -match $msol.ObjectId.Guid) { "Direct" }
									elseif ($_ -eq $null) { "Direct:NoGUID" }
									else { $(Get-MsolGroup -ObjectId $_ | Select-Object -ExpandProperty Displayname) }
								}
							}
						}
					}
					else
					{
						$lic = switch ($Plan)
						{
							"$($Company):ENTERPRISEPACK" { 'E3' }
							"$($Company):STANDARDPACK" { 'E1' }
							"$($Company):MFA_STANDALONE" { 'MFA' }
							"$($Company):POWER_BI_STANDARD" { 'PowerBi Free' }
							"$($Company):EXCHANGEENTERPRISE" { 'ExchangeOnlinePlan2' }
							"$($Company):EMS" { 'Enterprise Mobility Security' }
							"$($Company):FLOW_FREE" { 'Microsoft Flow Free' }
							"$($Company):POWERAPPS_INDIVIDUAL_USER" { 'PowerApps and Logic Flows' }
							"$($Company):MCOEV" { 'Phone System' }
							"$($Company):POWER_BI_PRO" { 'PowerBi Pro' }
							"$($Company):POWER_BI_ADDON" { 'Power BI for Office 365 Add-On' }
							"$($Company):POWER_BI_INDIVIDUAL_USER" { 'Power BI Individual User' }
							"$($Company):ENTERPRISEWITHSCAL" { 'Enterprise Plan E4' }
							"$($Company):PROJECTONLINE_PLAN_1" { 'Project Online' }
							"$($Company):PROJECTCLIENT" { 'Project Pro for Office 365' }
							"$($Company):VISIOCLIENT" { 'Visio Pro Online' }
							"$($Company):STREAM" { 'Microsoft Stream' }
							"$($Company):POWERAPPS_VIRAL" { 'Microsoft Power Apps & Flow' }
							"$($Company):PROJECTESSENTIALS" { 'Project Lite' }
							"$($Company):PROJECTPROFESSIONAL" { 'Project Professional' }
							"$($Company):SPZA_IW" { 'App Connect' }
							"$($Company):PBI_PREMIUM_P1_ADDON" { 'Power Bi Premium' }
							"$($Company):DYN365_ENTERPRISE_P1_IW" { 'Dynamics 365 P1 Trial for Information Workers' }
							"$($Company):WINDOWS_STORE" { 'Windows Store for Business' }
							"$($Company):DEVELOPERPACK" { 'Developer Pack' }
							default { "$_" }
						}
						
						$licobj = [PScustomobject]@{
							AccountSku = $lic
							Assignment = $null
						}
					}
					
					$prop.Add($licobj.AccountSku, $licobj.Assignment)
					
				}
				$prop.Add('Details', 'None')
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Enable-365MFA
{
<#
	.SYNOPSIS
		A brief description of the Enable-QHMFA function.
	
	.DESCRIPTION
		A detailed description of the Enable-QHMFA function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Enable-QHMFA -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$auth = New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement
					$auth.RelyingParty = "*"
					$auth.State = "Enabled"
					$auth.RememberDevicesNotIssuedBefore = (Get-Date)
					
					Set-MsolUser -UserPrincipalName $UPN -StrongAuthenticationRequirements $auth -ErrorAction Stop
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SUCCESS'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : Not a User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress		 = $UPN
					RecipientTypeDetails = 'ERROR'
					Status			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Enabling MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Enabling MFA' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Reset-365MFA
{
<#
	.SYNOPSIS
		A brief description of the Reset-QhMFASettings function.
	
	.DESCRIPTION
		A detailed description of the Reset-QhMFASettings function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Reset-QhMFASettings -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Parameter(Position = 1)]
		[switch]$ShowProgress
	)
	
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$mfa = @()
				$paramSetMsolUser = @{
					UserPrincipalName		    = $UPN
					StrongAuthenticationMethods = $mfa
					ErrorAction				    = 'Stop'
				}
				
				Set-MsolUser @paramSetMsolUser
				
				$prop = [Ordered]@{
					UserPrincipalName = $UPN
					MFAReset		  = 'Success'
					Details		      = 'None'
				}
			}
			catch
			{
				$prop = [Ordered]@{
					UserPrincipalName = $UPN
					MFAReset		  = 'Failed'
					Details		      = "Error : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Resetting MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
		Write-Progress -Activity 'Resetting MFA' -Completed
	}
}

function Disable-365MFA
{
<#
	.SYNOPSIS
		A brief description of the Disable-QHMFA function.
	
	.DESCRIPTION
		A detailed description of the Disable-QHMFA function.
	
	.PARAMETER UserPrincipalName
		A description of the UserPrincipalName parameter.
	
	.PARAMETER ShowProgress
		A description of the ShowProgress parameter.
	
	.EXAMPLE
		PS C:\> Disable-QHMFA -UserPrincipalName 'value1'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[Switch]$ShowProgress
	)
	begin
	{
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
		$i = 1
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$recipient = Get-Recipient $UPN -ErrorAction Stop
				if ($recipient.RecipientTypeDetails -match 'UserMailbox')
				{
					$auth = @()
					
					$paramSetMsolUser = @{
						UserPrincipalName			     = $UPN
						StrongAuthenticationRequirements = $auth
						ErrorAction					     = 'Stop'
					}
					
					Set-MsolUser @paramSetMsolUser
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SUCCESS'
						Details			     = 'None'
					}
				}
				else
				{
					$prop = [ordered]@{
						EmailAddress		 = $UPN
						RecipientTypeDetails = $recipient.RecipientTypeDetails
						Status			     = 'SKIPPED'
						Details			     = "SKIPPED : Not a User Mailbox"
					}
				}
			}
			catch
			{
				$prop = [ordered]@{
					EmailAddress		 = $UPN
					RecipientTypeDetails = 'ERROR'
					Status			     = 'FAILED'
					Details			     = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($ShowProgress)
				{
					if ($UserPrincipalName.count -gt 1)
					{
						$paramWriteProgress = @{
							Activity = 'Enabling MFA'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Enabling MFA' -Completed
		$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Get-365MFAStatus
{
<#
	.SYNOPSIS
		Gets MFA status of any user.
	
	.DESCRIPTION
		Gets MFA detailed status for any user like Default MFA method and other configured methods.
	
	.PARAMETER UserPrincipalName
		This is the User Principal Name of the user, sometimes it can be the Primary Email Address.
	
	.PARAMETER ShowProgress
		This will show progress of the command if processing bulk processing. It will not work when used with the pipeline.
	
	.EXAMPLE
		PS C:\> Get-365MFAStatus -UserPrincipalName 'user1@contoso.com'
	
	.EXAMPLE
		PS C:\> Get-365MFAStatus -UserPrincipalName 'user1@contoso.com','user1@contoso.com'
	
	.EXAMPLE
		PS C:\> Get-365MFAStatus -UserPrincipalName (Import-csv users.csv) -ShowProgress
	
	.EXAMPLE
		PS C:\> Get-MsolUser -All | Get-365MFAStatus
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   Position = 0,
				   HelpMessage = 'Enter UserPrincipal Name ?')]
		[ValidateNotNullOrEmpty()]
		[String[]]$UserPrincipalName,
		[switch]$ShowProgress
	)
	
	begin
	{
		$i = 1
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	}
	process
	{
		foreach ($UPN in $UserPrincipalName)
		{
			try
			{
				$msol = Get-MsolUser -UserPrincipalName $UPN -ErrorAction Stop
				if ($MSOL.StrongAuthenticationRequirements.State -eq $null)
				{
					$state = 'PotentiallyUnlicensed'
				}
				else
				{
					$state = $MSOL.StrongAuthenticationRequirements.State
				}
				
				if ($msol.StrongAuthenticationMethods.Count -eq 0)
				{
					$MFASetup = 'NotRegistered'
					$DefaultMethod = 'None'
					$OtherMethods = 'None'
					$details = 'None'
				}
				else
				{
					$MFASetup = 'Registered'
					$DefaultMethod = "$(($msol.StrongAuthenticationMethods.Where({ $_.'IsDefault' })).MethodType)"
					$OtherMethods = "$(($msol.StrongAuthenticationMethods.Where({ !($_.'IsDefault') })).MethodType -join (','))"
					$details = "None"
				}
				
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					MFAState		  = $state
					MFARegistration   = $MFASetup
					DefaultMFAMethod  = $DefaultMethod
					OtherMFAMethods	  = $OtherMethods
					Details		      = $details
				}
			}
			catch
			{
				$prop = [ordered]@{
					UserPrincipalName = $UPN
					MFAState		  = 'ERROR'
					MFARegistration   = 'ERROR'
					DefaultMFAMethod  = 'ERROR'
					OtherMFAMethods	  = 'ERROR'
					Details		      = "ERROR : $($_.Exception.Message)"
				}
			}
			finally
			{
				$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
				Write-Output $obj
				
				if ($UserPrincipalName.count -gt 1)
				{
					if ($ShowProgress)
					{
						$paramWriteProgress = @{
							Activity = 'Getting MFA License Status'
							Status   = "Processing [$i] of [$($UserPrincipalName.Count)] users"
							PercentComplete = (($i / $UserPrincipalName.Count) * 100)
							CurrentOperation = "Completed : [$UPN]"
						}
						Write-Progress @paramWriteProgress
					}
				}
				$i++
			}
		}
	}
	end
	{
		Write-Progress -Activity 'Getting MFA License Status' -Completed
		$Global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
	}
}

function Write-Log
{
<#
	.SYNOPSIS
		Logging Function
	
	.DESCRIPTION
		This function is used to create verbose logging
	
	.PARAMETER Type
		- INFO - informational logging messages
		- ERROR - Error logging messages
		- SUCCESS - Success logging messages
	
	.PARAMETER Message
		This is custom messages used with logging.
	
	.PARAMETER OnScreen
		This switch displays verbose messages on the screen.
	
	.PARAMETER Function
		This displays logging inforamtion about the functions called.
	
	.PARAMETER seperator
		A description of the seperator parameter.
	
	.EXAMPLE
		PS C:\> Write-QHLog -Type INFO -Message 'This is an information log' -Function 'foo-bar'
	
	.EXAMPLE
		PS C:\> Write-QHLog -Type ERROR -Message 'This is an Error log' -Function 'foo-bar'
	
	.EXAMPLE
		PS C:\> Write-QHLog -Type SUCCESS -Message 'This is a Success log' -Function 'foo-bar'
	
	.NOTES
		Additional information about the function.
#>
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[ValidateSet('INFO', 'ERROR', 'SUCCESS')]
		[string]$Type,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[String]$Message,
		[Parameter(Position = 2)]
		[switch]$OnScreen,
		[String]$Function = $($MyInvocation.InvocationName),
		[String]$seperator = '::'
	)
	begin
	{
		function seperator
		{
			param
			(
				[Parameter(Mandatory = $true)]
				[ValidateNotNullOrEmpty()]
				[String]$char
			)
			
			Write-Host -NoNewline " $char " -ForegroundColor Magenta
		}
	}
	process
	{
		$time = (Get-date).ToString('dd-MM-yyyy HH:mm:ss')
		
		$prop = [ordered]@{
			DateTime = $time
			Type	 = $Type
			Function = $Function
			Details  = $Message
		}
		
		if ($OnScreen)
		{
			switch ($Type)
			{
				'INFO' {
					$col = 'Yellow'
				}
				'ERROR' {
					$col = 'Red'
				}
				'SUCCESS' {
					$col = 'Green'
				}
				default
				{
					#<code>
				}
			}
			
			#$StringMsg = "$time :: $Type :: $message"
			Write-Host "$time" -ForegroundColor Gray -NoNewline
			seperator -char $seperator
			Write-Host "$Type" -ForegroundColor $col -NoNewline
			seperator -char $seperator
			Write-Host "$Function" -ForegroundColor White -NoNewline
			seperator -char $seperator
			Write-Host $Message -ForegroundColor Cyan
		}
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
		Write-Output $obj
	}
	end
	{
	}
}

function Generate-ComplexPassword
{
    <#
    .SYNOPSIS
    Password Generator
    
    .DESCRIPTION
    Password Generator tool to obtain any length and numbers of passwords, 
    adding desired number of special characters, quickly. 
    
    .PARAMETER PasswordLength
    Add a integer value for desired password length
    
    .PARAMETER SpecialCharCount
     Add a integer value for desired number of special characters
    
    .PARAMETER GenerateUserPW
    Enter as many named string or integer values 
    
    .EXAMPLE
    'John','Paul','George','Ringo' | New-ComplexPassword -PasswordLength 10 -SpecialCharCount 2
 
    1..5 | New-ComplexPassword -PasswordLength 16 -SpecialCharCount 5
    
    .NOTES
    #>
	
	[Cmdletbinding(DefaultParameterSetName = 'Single')]
	param (
		[Parameter(ParameterSetName = 'Single')]
		[Parameter(ParameterSetName = 'Multiple')]
		[Int]$PasswordLength,
		[Parameter(ParameterSetName = 'Single')]
		[Parameter(ParameterSetName = 'Multiple')]
		[int]$SpecialCharCount,
		[Parameter(ValueFromPipeline = $true,
				   ValueFromPipelineByPropertyName = $true,
				   ParameterSetName = 'Multiple')]
		[String[]]$GenerateUserPW
	)
	begin
	{
		# The System.Web namespaces contain types that enable browser/server communication
		Add-Type -AssemblyName System.Web
	}
	process
	{
		switch ($PsCmdlet.ParameterSetName)
		{
			'Single' {
				# GeneratePassword static method: Generates a random password of the specified length
				[System.Web.Security.Membership]::GeneratePassword($PasswordLength, $SpecialCharCount)
			}
			'Multiple' {
				$GenerateUserPW | ForEach-Object {
					# Custom Object to display results
					New-Object -TypeName PSObject -Property @{
						User	 = $_
						Password = [System.Web.Security.Membership]::GeneratePassword($PasswordLength, $SpecialCharCount)
					}
				}
			}
		}
	}
	end { }
}

function Split-Array
{
<#
	.SYNOPSIS
		Splits .
	
	.DESCRIPTION
		A detailed description of the Split-Array function.
	
	.PARAMETER inArray
		A description of the inArray parameter.
	
	.PARAMETER parts
		A description of the parts parameter.
	
	.PARAMETER size
		A description of the size parameter.
	
	.EXAMPLE
		PS C:\> Split-Array
	
	.NOTES
		Additional information about the function.
#>
	param (
		$inArray,
		[int]$parts,
		[int]$size
	)
	if ($parts)
	{
		$PartSize = [Math]::Ceiling($inArray.count / $parts)
	}
	if ($size)
	{
		$PartSize = $size
		$parts = [Math]::Ceiling($inArray.count / $size)
	}
	
	$outArray = New-Object 'System.Collections.Generic.List[psobject]'
	
	for ($i = 1; $i -le $parts; $i++)
	{
		$start = (($i - 1) * $PartSize)
		$end = (($i) * $PartSize) - 1
		if ($end -ge $inArray.count) { $end = $inArray.count - 1 }
		$outArray.Add(@($inArray[$start .. $end]))
	}
	return, $outArray
}