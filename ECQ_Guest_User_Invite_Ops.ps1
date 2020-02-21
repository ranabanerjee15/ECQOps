<#	
	.NOTES
	===========================================================================
	 Created with: 	SAPIEN Technologies, Inc., PowerShell Studio 2020 v5.7.172
	 Created on:   	19/02/2020 1:27 PM
	 Created by:   	Rana Banerjee
	 Organization: 	ECQ
	 Filename:     	ECQ_Guest_User_Invite_Ops.PS1
	===========================================================================
	.DESCRIPTION
		This Script will Invite new users as guest to Elections Tenant SharePoint.
#>

# *******************
#	Conditions
# *******************
#region Conditions

# %ForcePlatform% = x64
#Requires -Module ECQOps,AzureAD,ImportExcel

#endregion Conditions

# *******************
#	User Variables
# *******************
#region User and Email Variables

#Directory
$parent = 'D:\Office365\RB\Deprovops'

#Email Settings
$ReportRecipients = 'Rana.Banerjee@ecq.qld.gov.au'
$SmtpServer = 'Smtp.ecq.qld.gov.au'
$Subject = 'User Invite Process for Elections Tenant'
$from = 'Automation@ecq.qld.gov.au'

#endregion User Variables


# *******************
#	Script Variables
# *******************
#region Script Variables

#Credential clixml for ECQ tenant
$ECQAzCreds = $(Import-Clixml "$env:USERPROFILE\Cred\ECQAzCred.clixml")


#Credential clixml for Elections tenant
$ElectionsAzCreds = $(Import-Clixml "$env:USERPROFILE\Cred\ElectionsAzCred.clixml")

#Process current DateTime
$dateTime = (Get-date).ToString('dd-MM-yyyy-HH-mm-ss')

#Process Logfile Name and Path
$LogFileName = "GuestUser_Log_$($dateTime).csv"
$logfile = Join-Path -Path $parent -ChildPath $LogFileName

#Invited Users FileName and Path
$InvitedUsersFileName = "Invited_GuestUsers_$($dateTime).csv"
$InvitedUsersFile = Join-Path -Path $parent -ChildPath $InvitedUsersFileName

#Skipped Users FileName and Path
$SkippedUsersFileName = "Skipped_GuestUsers$($dateTime).csv"
$SkippedUsersFile = Join-Path -Path $parent -ChildPath $SkippedUsersFileName

#$deprovFileName = "Deproved_Report_$($dateTime).csv"
#$deprovedReport = Join-Path -Path $parent -ChildPath $deprovFileName
#$SummaryFileName = "HHS_DeProv_Summary_Report.csv"
#$SummaryReport = Join-Path -Path $parent -ChildPath $SummaryFileName

#Excel Report File
$ExcelReportFileName = "InvitedUsersReport_$($dateTime).xlsx"
$ExcelReport = Join-Path -Path $parent -ChildPath $ExcelReportFileName

#This will show Progress - Yet to impliment
$showProgress = $true

 
# DO NOT CHANGE THIS. 
# THIS IS SET DYNAMICALLY BY THE PROCESS.
# AT THE BEGNING ITS SET TO TRUE AND CHANGED TO FALSE ON TERMINATING ERROR
$Script:Proceed = $true

#endregion Script Variables


# *******************
#	Required Modules
# *******************
#region Required Modules

try
{
	$msg = "Trying to Import Required modules"
	Write-Log -Type INFO -Message $msg -Function $MyInvocation.InvocationName -OnScreen #| Export-Csv $logfile -Append -NoTypeInformation
	
	Import-Module AzureAD -ErrorAction Stop -WarningAction SilentlyContinue
	Import-Module ECQOps -ErrorAction Stop -WarningAction SilentlyContinue
	Import-Module ImportExcel -ErrorAction Stop -WarningAction SilentlyContinue
	
	$msg = "Successfully imported Required modules"
	Write-Log -Type SUCCESS -Message $msg -Function $MyInvocation.InvocationName -OnScreen #| Export-Csv $logfile -Append -NoTypeInformation
	
}
catch
{
	$msg = "$($_.Exception.Message). Terminating Process"
	Write-Log -Type ERROR -Message $msg -Function $MyInvocation.InvocationName -OnScreen # | Export-Csv $logfile -Append -NoTypeInformation
	$script:Proceed = $false
}

#endregion Required Modules


# *******************
#	Script Functions
# *******************
#region Functions

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
		PS C:\> Write-Log -Type INFO -Message 'This is an information log' -Function 'foo-bar'
	
	.EXAMPLE
		PS C:\> Write-Log -Type ERROR -Message 'This is an Error log' -Function 'foo-bar'
	
	.EXAMPLE
		PS C:\> Write-Log -Type SUCCESS -Message 'This is a Success log' -Function 'foo-bar'
	
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

function Connect-x365AzureAD
{
<#
	.SYNOPSIS
		Connects to Azure AD Service
	
	.DESCRIPTION
		Connects to Azure AD Service. This uses PowerShell Module 'AZUREAD'. Please make sure you have installed this Module before running this command
	
	.PARAMETER Credential
		This is the credential which you will use to connect to Azure AD Service.
	
	.EXAMPLE
		PS C:\> Connect-x365AzureAD -Credential (Get-Credential)
	
	.EXAMPLE
		PS C:\> Connect-x365AzureAD -Credential 'Adminuser@domain.onmicrosoft.com'
	
	.EXAMPLE
		PS C:\> Connect-x365AzureAD -Credential $cred
	
	.NOTES
		In the above example $cred can be a variable which has your credentials stored.
#>
	
	param
	(
		[System.Management.Automation.Credential()]
		[ValidateNotNull()]
		[System.Management.Automation.PSCredential]$Credential = [System.Management.Automation.PSCredential]::Empty
	)
	try
	{
		$msg = "Trying to Connect to Azure AD Service"
		Write-Log -Type INFO -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		
		$FormatEnumerationLimit = -1
		#Write-Host "INFO : Trying to Connect to Office 365" -ForegroundColor Cyan
		
		if ($credential -eq $null)
		{
			$credential = Get-Credential -Message "Enter your Credentials" -ErrorAction Stop
		}
		Connect-AzureAd -Credential $Credential -ErrorAction Stop
		#Write-Host "SUCCESS : Successfully Connected to Office 365" -ForegroundColor Green
		$msg = "Successfully connected to Azure AD Service"
		Write-Log -Type SUCCESS -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		
	}
	catch
	{
		$msg = "$($_.Exception.Message). Terminating Process"
		Write-Log -Type ERROR -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		$script:Proceed = $false
		break
		#exit
	}
}

function Disconnect-x365AzureAD
{
<#
	.SYNOPSIS
		Disconnects from Azure AD Service

	.DESCRIPTION
		Disconnects from Azure AD Service. This uses PowerShell Module 'AZUREAD'. Please make sure you have installed this Module before running this command

	.EXAMPLE
		PS C:\> Disconnect-x365AzureAD

	.NOTES
		
#>
	[CmdletBinding()]
	param ()
	
	try
	{
		$msg = "Trying to Disconnect From Azure AD Service"
		Write-Log -Type INFO -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		
		$FormatEnumerationLimit = -1
		#Write-Host "INFO : Trying to Connect to Office 365" -ForegroundColor Cyan
		
		Disconnect-AzureAd -ErrorAction Stop
		#Write-Host "SUCCESS : Successfully Connected to Office 365" -ForegroundColor Green
		$msg = "Successfully connected to Azure AD Service"
		Write-Log -Type SUCCESS -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		
	}
	catch
	{
		$msg = "$($_.Exception.Message). Terminating Process"
		Write-Log -Type ERROR -Message $msg -Function $MyInvocation.InvocationName -OnScreen | Export-Csv $logfile -Append -NoTypeInformation
		$script:Proceed = $false
		break
		#exit
	}
}

function New-ECQUserInvition
{
<#
	.SYNOPSIS
		This command will invite a new ECQ Tenant User to Elections tenant
	
	.DESCRIPTION
		This command will send new guest invitation to ECQ Tennant user for accessing Sharepoint resources in Elections Tennant 
	
	.PARAMETER DisplayName
		This will be the Diplay Name of the invited user in ECQ Tennant
	
	.PARAMETER UserPrincipalName
		This will be the UserPrincipalName of the invited user in ECQ Tennant
	
	.EXAMPLE
		PS C:\> New-ECQUserInvition -DisplayName 'Joe Bloke' -EmailAddress 'Joe.Bloke@ecq.qld.gov.au'
	
	.NOTES
		Additional information about the function.
#>
	
	[CmdletBinding()]
	param
	(
		[Parameter(Mandatory = $true,
				   Position = 0)]
		[ValidateNotNullOrEmpty()]
		[String]$DisplayName,
		[Parameter(Mandatory = $true,
				   Position = 1)]
		[ValidateNotNullOrEmpty()]
		[String]$EmailAddress
	)
	
	#TODO: Place script here
	
	$null = $obj,$prop
	
	Try
	{
		$paramNewAzureADMSInvitation = @{
			InvitedUserDisplayName  = $DisplayName
			InvitedUserEmailAddress = $EmailAddress
			SendInvitationMessage   = $false
			InviteRedirectUrl	    = "https://ecqgovelec.sharepoint.com"
			InvitedUserType		    = 'member'
			ErrorAction			    = 'Stop'
		}
		
		New-AzureADMSInvitation @paramNewAzureADMSInvitation | Out-Null
		
		$prop = [Ordered]@{
			InvitedUserDisplayname = $DisplayName
			InvitedUserEmailAddress = $EmailAddress
			Status				    = 'Success'
			Details = 'None'
		}
	}
	Catch
	{
		$prop = [Ordered]@{
			InvitedUserDisplayname  = $DisplayName
			InvitedUserEmailAddress = $EmailAddress
			Status				    = 'Failed'
			Details				    = "Error : $($_.Exception.Message)"
		}
	}
	Finally
	{
		$obj = New-Object -TypeName System.Management.Automation.PSObject -Property $prop
		Write-Output $obj
	}
}

#endregion Functions


# *******************
#	Main Function
# *******************
#region Function Main

#Starting the Main Function

function Main
{
	
	#Will execute process until there is a terminating error
	while ($script:Proceed)
	{
		#Sets the working Directory to parent 
		Set-Location $parent | Out-Null
		$msg = "Starting Process Guest User access on Elections Tenant"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
		
		$msg = "Connecting to ECQ Tennant"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
		
		
		#Connects to ECQ Azure Tenant
		Connect-x365AzureAD -Credential $ECQAzCreds
		
		$msg = "Retriving Users from ECQ Tennant"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
		
		#Retrives all the ECQ users with their DisplayName and UserPrincipalName
		[Array]$ECQUsers = Get-AzureADUser -All $true -ErrorAction SilentlyContinue |
		Select-Object Displayname,UserPrincipalName
		
		
		#Only processes the ECQ users if it retrives ECQ users
		if ($ECQUsers.Count -eq 0)
		{
			# Terminates the process as nothing to process
			$msg = "No Users to process or could not retrive any users. Terminating the Process"
			Write-Log -Type ERROR -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
			Export-Csv $logfile -Append -NoTypeInformation
			$script:Proceed = $false
		}
		Else
		{
			$msg = "Retrived $($ECQUsers.Count) ECQ Users"
			Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
			Export-Csv $logfile -Append -NoTypeInformation
		}
		
		#Disconnects from the ECQ Tenant
		Disconnect-x365AzureAD
		
		
		$msg = "Connecting to Elections Tennant"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
		
		#Connects to Elections Azure Tenant
		Connect-x365AzureAD -Credential $ElectionsAzCreds
		
		$msg = "Retriving Guest Users from Elections Tennant"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
		
		#Retrives all GUEST users from Elections Tenant with their UserPrincipalName
		[Array]$ElectionsGuestUsers = Get-AzureADUser -Filter "UserType eq 'Guest'" -ErrorAction SilentlyContinue |
		Select-Object -ExpandProperty UserPrincipalName
		
		if ($ElectionsGuestUsers.Count -eq 0)
		{
			$msg = "No Guest Users in Elections Tennant to compare or could not retrive them. Terminating the Process"
			Write-Log -Type ERROR -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
			Export-Csv $logfile -Append -NoTypeInformation
			$script:Proceed = $false
		}
		Else
		{
			$msg = "Retrived $($ElectionsGuestUsers.Count) Election Users to compare against"
			Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
			Export-Csv $logfile -Append -NoTypeInformation
		}
		
		
		# Create a Generic List to hold the user PsObject
		#$Processed = New-Object 'System.Collections.Generic.List[psobject]'
		
		#Compare Users from both the Tenants
		$i = 1
		foreach ($ECQUser in $ECQUsers)
		{
			$null = $usrobj
			
			$msg = "Checking if $($ECQUser.UserPrincipalName) exists in Elections Tenant"
			Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
			Export-Csv $logfile -Append -NoTypeInformation
			
			if($ElectionsGuestUsers -inotcontains $ECQUser.UserPrincipalName)
			{
				#Processing needed
				
				$msg = "User $($ECQUser.UserPrincipalName) Does not exists in Elections Tenant, Trying to send invite now"
								
				Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
				Export-Csv $logfile -Append -NoTypeInformation
				
				$paramNewECQUserInvition = @{
					DisplayName = $ECQUser.UserPrincipalName
					EmailAddress = $ECQUser.UserPrincipalName
				}
				
				$usrobj = New-ECQUserInvition @paramNewECQUserInvition
				
				$usrobj | Export-Csv $InvitedUsersFile -Append -NoTypeInformation
			
			}
			Else
			{
				# No processing needed
				
				$msg = "User $($ECQUser.UserPrincipalName) already exists in Elections Tenant, No Action taken"
				
				Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
				Export-Csv $logfile -Append -NoTypeInformation
				$ECQUser | Export-Csv $SkippedUsersFile -Append -NoTypeInformation
			}
			
			if ($showProgress)
			{
				$paramWriteProgress = @{
					Activity = 'Processing Users for Guest Invite in Elections Tenant'
					Status   = "Processing [$i] of [$($ECQUsers.Count)] users"
					PercentComplete = (($i / $ECQUsers.Count) * 100)
					CurrentOperation = "Completed : [$($ECQUser.UserPrincipalName)]"
				}
				
				Write-Progress @paramWriteProgress
			}
			
		}
		
		$msg = "Processed total [$($ECQUsers.count)] out of which Invited [$($Proceed.count)] users."
		
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
		Export-Csv $logfile -Append -NoTypeInformation
	}
	
	#Process for Users complete
	
	#Check for generated log files for report.
	
	$msg = "Attempting to generate Excel report"
	Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen |
	Export-Csv $logfile -Append -NoTypeInformation
	
	if (Test-Path $InvitedUsersFile)
	{
		$text1 = New-ConditionalText 'Success' -ConditionalTextColor White -BackgroundColor '#FD625E' -Range C:C
		$text2 = New-ConditionalText 'Failed' -ConditionalTextColor White -BackgroundColor '#00b359' -Range C:C
		#$text3 = New-ConditionalText 'PotentiallyUnlicensed' -ConditionalTextColor White -BackgroundColor '#3599B8'
		$text4 = New-ConditionalText 'ERROR' -ConditionalTextColor Black -BackgroundColor '#F2C80F' -Range D:D
		
		$paramExportExcel = @{
			Path		    = $ExcelReport
			Show		    = $false
			AutoSize	    = $true
			TableName	    = 'InvitedUsers'
			TableStyle	    = 'Medium20'
			FreezeTopRow    = $true
			BoldTopRow	    = $true
			WorkSheetname   = 'Invited Users'
			ConditionalText = $text4, $text1, $text2
			IncludePivotTable = $true
			#PivotRows	    = 'Status'
			#PivotData	    = 'Status'
			#IncludePivotChart = $true
			#ChartType	    = 'PieExploded3D'
		}
		
		Import-Csv $InvitedUsersFile | Export-Excel @paramExportExcel
	}
	
	if (Test-Path $SkippedUsersFile)
	{
		
	}
	
	If (Test-Path $logfile)
	{
		$text1 = New-ConditionalText 'ERROR' -ConditionalTextColor White -BackgroundColor '#FD625E' -Range B:B
		$text2 = New-ConditionalText 'Success' -ConditionalTextColor White -BackgroundColor '#00b359' -Range B:B
		$text3 = New-ConditionalText 'INFO' -ConditionalTextColor White -BackgroundColor '#3599B8' -Range B:B
		# $text4 = New-ConditionalText 'ERROR' -ConditionalTextColor Black -BackgroundColor '#F2C80F'
		
		$paramExportExcel = @{
			Path		    = $ExcelReport
			Show		    = $false
			AutoSize	    = $true
			TableName	    = 'TableLogs'
			TableStyle	    = 'Light8'
			FreezeTopRow    = $true
			BoldTopRow	    = $true
			WorkSheetname   = 'Logs'
			ConditionalText = $text3, $text1, $text2
		}
		Import-Csv $logfile | Export-Excel @paramExportExcel
		
	}
	
	# Attempting to Email Reports and logs
	
	$paramMail = @{ }
	$paramMail.Subject = $Subject
	$paramMail.To = $ReportRecipients
	$paramMail.From = $from
	$paramMail.SmtpServer = $SmtpServer
	$paramMail.BodyAsHtml = $true
	$paramMail.ErrorAction = 'Stop'
	
	$attachments = @()
	
	$logfile,$InvitedUsersFile,$SkippedUsersFile | ForEach-Object { if (Test-Path $_) { $attachments += $_ } }
	
	if ($attachments.count -gt 0)
	{
		$paramMail.Attachments = $attachments
	}
	
	try
	{
		$msg = "Trying to Email Report to defined recipients"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen | Export-Csv $LogFile -Append -NoTypeInformation
		
		Send-MailMessage @paramMail
		
		$msg = "Report successfully Emailed Reports to the defined Recipients"
		Write-Log -Type SUCCESS -Function $MyInvocation.InvocationName -Message $msg -OnScreen | Export-Csv $LogFile -Append -NoTypeInformation
		$msg = "Completed the Entire process, Exiting the Process"
		Write-Log -Type INFO -Function $MyInvocation.InvocationName -Message $msg -OnScreen | Export-Csv $LogFile -Append -NoTypeInformation
		
	}
	catch
	{
		$msg = "$($_.Exception.Message)"
		Write-Log -Type ERROR -Function $MyInvocation.InvocationName -Message $msg -OnScreen | Export-Csv $LogFile -Append -NoTypeInformation
		
	}
}

#endregion Function Main


# *******************
#	Execution
# *******************
#region Execution

Main

#endregion Execution

