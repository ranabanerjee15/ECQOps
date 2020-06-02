<#
	.SYNOPSIS
		A brief description of the Generate-O365LicenseReport.ps1 file.
	
	.DESCRIPTION
		This Script will generate office 365 License Report in Excel
	
	.PARAMETER ExcelReport
		A description of the ExcelReport parameter.
	
	.NOTES
		===========================================================================
		Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.178
		Created on:   	29/05/2020 4:32 PM
		Created by:   	Rana Banerjee
		Organization: 	ECQ
		Filename:		Generate-O365LicenseReport
		===========================================================================
#>
[CmdletBinding()]
param
(
	[Parameter(Position = 1)]
	[string]$ExcelReport,
	[switch]$Show
)
BEGIN
{
	$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Stop
	#region Functions
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
										"$($Company):ENTERPRISEPACK" { 'E3' }
										"$($Company):STANDARDPACK" { 'E1' }
										"$($Company):MFA_STANDALONE" { 'MFA' }
										"$($Company):POWER_BI_STANDARD" { 'PowerBi Free' }
										"$($Company):EXCHANGEENTERPRISE" { 'Exchange Online Plan2' }
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
										"$($Company):THREAT_INTELLIGENCE" { 'Office 365 Threat Intelligence' }
										"$($Company):OFFICESUBSCRIPTION" { 'Office 365 ProPlus' }
										"$($Company):MCOMEETADV" { 'Skype for Business PSTN Conferencing' }
										"$($Company):SPE_E3" { 'Secure Productive Enterprise E3' }
										"$($Company):ATP_ENTERPRISE" { 'Exchange Online ATP' }
										"$($Company):EXCHANGESTANDARD" { 'Exchange Online Plan 1' }
										"$($Company):SPE_E5" { 'Microsoft 365 E5' }
										"$($Company):RMSBASIC" { 'RMS Basic' }
										"$($Company):VISIOONLINE_PLAN1" { 'Visio Online Plan 1' }
										default { "$_" }
									}
								}
							},
														   @{
								n								  = 'Assignment'; e = {
									$lic.GroupsAssigningLicense.Guid | ForEach-Object {
										if ($_ -match $msol.ObjectId.Guid) { "Licensed" }
										elseif ($_ -eq $null) { "Licensed" }
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
								"$($Company):EXCHANGEENTERPRISE" { 'Exchange Online Plan 2' }
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
								"$($Company):THREAT_INTELLIGENCE" { 'Office 365 Threat Intelligence' }
								"$($Company):OFFICESUBSCRIPTION" { 'Office 365 ProPlus' }
								"$($Company):MCOMEETADV" { 'Skype for Business PSTN Conferencing' }
								"$($Company):SPE_E3" { 'Secure Productive Enterprise E3' }
								"$($Company):ATP_ENTERPRISE" { 'Exchange Online ATP' }
								"$($Company):EXCHANGESTANDARD" { 'Exchange Online Plan 1' }
								"$($Company):SPE_E5" { 'Microsoft 365 E5' }
								"$($Company):RMSBASIC" { 'RMS Basic' }
								"$($Company):VISIOONLINE_PLAN1" { 'Visio Online Plan 1' }
								default { "$_" }
							}
							
							$licobj = [PScustomobject]@{
								AccountSku = $lic
								Assignment = 'None'
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
	function Get-365LicenseStatusSummary
	{
		Try
		{
			$AllSkus = Get-MsolAccountSku -ErrorAction Stop
			$Company = $AllSkus[0].AccountName
			$AllSkus |
			Select-Object @{
				n = "Licenses"; e = {
					switch ($_.AccountSkuid)
					{
						"$($Company):ENTERPRISEPACK" { 'E3' }
						"$($Company):STANDARDPACK" { 'E1' }
						"$($Company):MFA_STANDALONE" { 'MFA' }
						"$($Company):POWER_BI_STANDARD" { 'PowerBi Free' }
						"$($Company):EXCHANGEENTERPRISE" { 'Exchange Online Plan2' }
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
						"$($Company):THREAT_INTELLIGENCE" { 'Office 365 Threat Intelligence' }
						"$($Company):OFFICESUBSCRIPTION" { 'Office 365 ProPlus' }
						"$($Company):MCOMEETADV" { 'Skype for Business PSTN Conferencing' }
						"$($Company):SPE_E3" { 'Secure Productive Enterprise E3' }
						"$($Company):ATP_ENTERPRISE" { 'Exchange Online ATP' }
						"$($Company):EXCHANGESTANDARD" { 'Exchange Online Plan 1' }
						"$($Company):SPE_E5" { 'Microsoft 365 E5' }
						"$($Company):RMSBASIC" { 'RMS Basic' }
						"$($Company):VISIOONLINE_PLAN1" { 'Visio Online Plan 1' }
						default { "$_" }
					}
				}
			},
			@{ n = 'TotalLicenses'; e = { $_.ActiveUnits }},
			@{ n = 'UsedLicenses'; e = { $_.ConsumedUnits }},
			@{ n = 'RemainingLicenses'; e = { $_.ActiveUnits - $_.ConsumedUnits }}
		}
		Catch
		{
			Write-Warning "ERROR : $($_.Exception.Message)"
		}
	}
	Function Get-ScriptLocation
	{
		# If using ISE
		if ($psISE)
		{
			$ScriptPath = Split-Path -Parent $psISE.CurrentFile.FullPath
			# If Using PowerShell 3 or greater
		}
		elseif ($PSVersionTable.PSVersion.Major -gt 3)
		{
			$ScriptPath = $PSScriptRoot
			# If using PowerShell 2 or lower
		}
		else
		{
			$ScriptPath = split-path -parent $MyInvocation.MyCommand.Path
		}
		
		Write-Output $ScriptPath
		#Write-host "$ScriptPath" -ForegroundColor Green
	}
	#endregion Functions
	
	if ('' -eq $ExcelReport)
	{
		$ExcelReport = "$(Get-ScriptLocation)\O365LicenseReport.xlsx"
	}
	
	Try
	{
		Import-Module ImportExcel -ErrorAction Stop
	}
	Catch
	{
		Write-Warning "ERROR:$($_.Exception.Message). Please make sure you have installed the 'ImportExcel' module and have the right permissions to load it. To install this module please run 'Install-Module -Name ImportExcel -Scope CurrentUser -AllowClobber -Force -Confirm:$false'. This will install the module in the user scope. Terminating process."
		break
	}
	
	$session = Get-MsolAccountSku -ErrorAction SilentlyContinue
	if ($null -eq $session)
	{
		Try
		{
			Connect-MsolService
		}
		Catch
		{
			Write-Warning "ERROR:$($_.Exception.Message). Terminating Process."
			Break
		}
	}
	
}
PROCESS
{
	$LicensedUsers = Get-MsolUser -All | Sort-Object | Where-Object { $_.IsLicensed } | Select-Object -ExpandProperty UserPrincipalName
	$LicenseResultRaw = Get-365LicenseStatus -UserPrincipalName $LicensedUsers -ShowProgress
	$summary365License = Get-365LicenseStatusSummary | Sort-Object -Property UsedLicenses -Descending
	$text1 = New-ConditionalText 'Licensed' -ConditionalTextColor White -BackgroundColor ([System.Drawing.Color] '#00b359')
	#$text2 = New-ConditionalText 'Direct:NoGUID' -ConditionalTextColor White -BackgroundColor ([System.Drawing.Color] '#00b359')
	$text3 = New-ConditionalText 'ERROR' -ConditionalTextColor Black -BackgroundColor ([System.Drawing.Color] '#F2C80F')
	
	$paramExportExcel1 = @{
		Path		    = $ExcelReport
		Show		    = $false
		AutoSize	    = $true
		TableName	    = 'TableECQO365Licenses'
		FreezeTopRow    = $true
		BoldTopRow	    = $true
		WorkSheetname   = 'O365LicenseReport'
		ConditionalText = $text1, $text3
		TableStyle	    = 'Medium16'
	}
	
	$LicenseResultRaw | Sort-Object -Property UserPrincipalName | Export-Excel @paramExportExcel1
	
	$paramExportExcel2 = @{
		Path		    = $ExcelReport
		Show		    = $false
		AutoSize	    = $true
		TableName	    = 'TableECQO365LicenseSummary'
		FreezeTopRow    = $true
		BoldTopRow	    = $true
		WorkSheetname   = 'O365LicenseSummary'
		TableStyle	    = 'Medium17'
	}
	$summary365License | Export-Excel @paramExportExcel2
}
END
{
	if ($Show)
	{
		Invoke-Item $ExcelReport
	}
	$global:ErrorActionPreference = [System.Management.Automation.ActionPreference]::Continue
}
#TODO: Place script here

