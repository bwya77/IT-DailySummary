
<#	
	.NOTES
	===========================================================================
	 Created on: 		12/17/2018
	 Last Modified on:  1/2/2019 7:37 PM
	 Created by:   		Bradley Wyatt; bwyatt@psmpartners.com
	 Version: 	    	1.3.2
	 Modules:			ReportHTML, ActiveDirectory
	 Notes: This script will create a massive Active Directory overview report for you and your team. I have set it up as a daily scheduled task for our team:
		- Ability to filter out users, exclude samaccountname's of service accounts
		- Export encrypted PScredential object to send emails, will use it on each run so you do not need to manually enter credentials each time
		- Full detailed logging 
		- Gater users with Passwords expiring soon
		- Send the end user notifications on their password expiring soon and how to change it
			- It will send every day until they change it
			- Explain how to change their password
			- Include password complextiy requirements
		- Gather accounts that are expiring soon
		- Gather locked out accounts
		- Gather newly created accounts
		- Gather inactive accounts (accounts that have not logged on in X amount of days)
		- Gather accounts with passwords set to never expire
		- Gather accounts with no manager set
		- Send a summary email with HTML formatting of the data
		- Create an interactive HTML report with the data and pie graphs, attatch to summary email
		- Send a basic summary message to a Teams channel
	===========================================================================
	.DESCRIPTION
		This script will create a massive Active Directory overview report for you and your IT department. It will be formatted in HTML in the body of the email but also attatch an interactive HTML report to the email.
		It can also notify users with passwords expiring soon about the expiration and directions on how to reset it. 

		I set this up as a daily scheduled task 
#>

############################################################################################################
#                                                                                                          #
#							                 VARIABLES	        										   #
#                                                                                                          #
############################################################################################################

#The logo that will be on the right side of the HTML report
$RightLogo = "https://www.thelazyadministrator.com/wp-content/uploads/2018/08/ewwewewe.png"
#The logo that will be on the left side of the HTML report
$CompanyLogo = "https://www.psmpartners.com/wp-content/uploads/2017/10/porcaro-stolarek-mete.png"

#Tells the script whether to gather users that are currently locked out or not
$GatherLockedOutUsers = $true

#Tells the script whether to gather users with passwords expiring 'soon'(set the # of days below) or not
$GatherPasswordExpiringUsers = $true
#If enabled the script will notify users that their password is expiring in less that $expireindays and also include password reset instruction and password requirements 
$SendEndUserPasswordExpirationNoticeEmails = $true
#Set the days for password expiration gather. Ex: Get all users whos password expires in $expireindays days or less. Anything greater than $expireindays days will be disregarded
$expireindays = 7

#Tells the script whether to gather users that are currently expiring or not
$GatherExpiringAccounts = $true
#Get all accounts that are set to expire in X days and below
$ExpiringAccountDays = 7

#Tells the script whether to gather users that are newly created or not
$GatherNewlyCreatedUsers = $true
#Sets the days for what is considered 'newly created' EX: get all users that have been created within the last $UserCreatedDays days
$UserCreatedDays = 2

#Tells the script whether to gather users that have passwords set to never expire or not
$GatherUsersWithPasswordsNeverSetToExpire = $true

#Tells the script whether to gather users that are considered inactive or not
$GatherInactiveUsers = $true
#Sets the amount of days a user is considered inactive. EX: users that have not logged into the system in $InactiveDays days or more
$InactiveDays = 90

#Tells the script whether to gather users that do not have a manager set in Active Directory or not
$GatherUsersWithoutManagers = $true

#Enables a summary email notification (you will want this enabled unless you only want a basic Teams Message)
$SendSummaryEmailNotification = $True
#Sets the recipient for the summary emails
$SummaryEmailAddress = "Brad@TheLazyAdministrator.com"

#Enabled a basic summary message to be sent to teams via Webhook
$SendSummaryTeamMessage = $false
#Sets the Teams webhook url
$SummaryTeamWebhookURL = "https://outlook.office.com/webhook/8348ef8d-c544-4dee-9855-b4de08dsf87sd@6988798-bd25-4817-8fd6-98sdf78sdfc1/IncomingWebhook/8189d8f9s8d6cbd81b446c16/5f2b8c07-e613-4e86-bcf6-589d9sd7ad138"

#If you did not want pie graphs in the bottom of the HTML summary email set this to false
$DisableGraphs = $False

#Variables for Sending users notification emails if their passwords are expiring soon and summary email
#SMTP host
$SMTPHost = "smtp.office365.com"
#Who the message will be sent from, you will need these credentials or use credentials of an account that has permission to send as this account
$FromEmail = "Brad@TheLazyAdministrator.com"

#Program Directory path
$DirPath = "C:\Automation\Master_DailySummary"

#Exclude accounts, good for service accounts
$ExcludeUsers = $True
#Where there exclude list will be, fill it with SamAccountName's'
$ExcludeList = "$DirPath\excludeSamAccountName.txt"

############################################################################################################

############################################################################################################
#                                                                                                          #
#							           BEGIN OF SCRIPT	        										   #
#                                                                                                          #
############################################################################################################

#Counters
$Int_LessThan = 0
$Int_Today = 0
$Int_ExpiringAccounts = 0
$Int_LockedOut = 0
$Int_InactiveUsers = 0
$Int_PWNeverExpires = 0
$Int_NewAccounts = 0
$Int_NoEmail = 0
$Int_NoManager = 0

# Creating head style
$head = @"

  <style> 

  h1 {

  text-align:center;

  border-bottom:1px solid #666666;

  color:blue;

  }

  TABLE {

  TABLE-LAYOUT:  fixed; 

  FONT-SIZE:  100%; 

  WIDTH:  100%;

  BORDER: 1px  solid black;

  border-collapse: collapse;

  }

  * {

  margin:0

}
              .pageholder {

  margin:  0px auto;

  }

  

  td {

  VERTICAL-ALIGN:  TOP; 

  FONT-FAMILY:  Tahoma;

  WORD-WRAP:  break-word; 

  BORDER: 1px  solid black;          

  }

  

  th {

  VERTICAL-ALIGN:  TOP; 

  COLOR:  #018AC0; 

  TEXT-ALIGN:  left;

  background-color:#00aaff;

  color:#f9f9f9;

  BORDER: 1px  solid black;

  }

  body {

  text-align:left;

  font-smoothing:always;

  width:100%;

  }        

  .even {  background-color: #dddddd; }

  .odd { background-color:  #ffffff; }          

  </style>

"@



#If the exclude list is not present, and $ExcludeUsers is set to True, then create the flat file
If (((Test-Path -Path $ExcludeList) -eq $false) -and ($ExcludeUsers -eq $true))
{
	New-Item -ItemType file $ExcludeList -Force
}

(get-date -Format "dd/MM/yyyy hh:mm:ss") + (": NEW RUN") | Out-File ($DirPath + "\" + "Log.txt") -Append -ErrorAction SilentlyContinue


(get-date -Format hh:mm:ss) + (": Importing Active Directory module") | Out-File ($DirPath + "\" + "Log.txt") -Append
Write-Host "Importing ActiveDirectory Module"
Import-Module ActiveDirectory


(get-date -Format hh:mm:ss) + (": Checking for filter users") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($ExcludeUsers -eq $true)
{
	(get-date -Format hh:mm:ss) + (": Filter users enabled, getting samaccountname object from $ExcludeList") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$exclude = Get-Content $ExcludeList
	If ($Null -ne $exclude)
	{
		$filter = [scriptblock]::create(($exclude | ForEach-Object { "(SamAccountName -notlike '*$_*')" }) -join ' -and ')
		Write-Host "Getting filtered users..."
		$Users = Get-ADUser -filter $filter -properties *
	}
	Else
	{
		$Users = Get-Aduser -Filter * -properties *
		
	}
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Filter users off, getting all users") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Getting all Users..."
	(get-date -Format hh:mm:ss) + (": Getting all users") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Users = Get-Aduser -Filter * -properties *
}


(get-date -Format hh:mm:ss) + (": Checking for ReportHTML Module") | Out-File ($DirPath + "\" + "Log.txt") -Append
$Mod = Get-Module -ListAvailable -Name "ReportHTML"

(get-date -Format hh:mm:ss) + (": Checking for module result") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($null -eq $Mod)
{
	(get-date -Format hh:mm:ss) + (": Module not found") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "ReportHTML Module is not present, attempting to install it"
	(get-date -Format hh:mm:ss) + (": Installing module") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Try
	{
		Install-Module -Name ReportHTML -Force
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
	(get-date -Format hh:mm:ss) + (": Importing ReportHTML Module") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Try
	{
		Import-Module ReportHTML -ErrorAction SilentlyContinue
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
}

(get-date -Format hh:mm:ss) + (": Checking for program folder path") | Out-File ($DirPath + "\" + "Log.txt") -Append
Write-Host "Checking for program folder path..."
#Check if program dir is present
$DirPathCheck = Test-Path -Path $DirPath
If (!($DirPathCheck))
{
	(get-date -Format hh:mm:ss) + (": Program folder  not found") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Try
	{
		Write-Host "Folder path not found, creating folder"
		#If not present then create the dir
		(get-date -Format hh:mm:ss) + (": Creating program folder") | Out-File ($DirPath + "\" + "Log.txt") -Append
		New-Item -ItemType Directory $DirPath -Force
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
}
Else
{
	(get-date -Format hh:mm:ss) + (": Program foldert is present") | Out-File ($DirPath + "\" + "Log.txt") -Append
}


#CredObj path
$CredObj = ($DirPath + "\" + "EmailExpiry.cred")
#Check if CredObj is present
(get-date -Format hh:mm:ss) + (": Checking for cred object") | Out-File ($DirPath + "\" + "Log.txt") -Append
Write-Host "Checking for valid cred object..." -ForegroundColor Yellow
$CredObjCheck = Test-Path -Path $CredObj
If (!($CredObjCheck))
{
	(get-date -Format hh:mm:ss) + (": Cred object not found") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Credential object not found, creating" -ForegroundColor Red
	(get-date -Format hh:mm:ss) + (": Importing Cred Object") | Out-File ($DirPath + "\" + "Log.txt") -Append
	#If not present get office 365 cred to save and store
	(get-date -Format hh:mm:ss) + (": Prompting for credentials") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Credential = Get-Credential -Message "Please enter your Office 365 credential that you will use to send e-mail from $FromEmail. If you are not using the account $FromEmail make sure this account has 'Send As' rights on $FromEmail."
	#Export cred obj
	(get-date -Format hh:mm:ss) + (": Exporting Credentials") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Try
	{
		$Credential | Export-CliXml -Path $CredObj
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Cred object found") | Out-File ($DirPath + "\" + "Log.txt") -Append
}
(get-date -Format hh:mm:ss) + (": Importing the cred object") | Out-File ($DirPath + "\" + "Log.txt") -Append
Write-Host "Importing Cred object..." -ForegroundColor Yellow
Try
{
	$Cred = (Import-CliXml -Path $CredObj)
}
Catch
{
	$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
}


###################################################################################################

$ExpiringUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Expiring Accounts") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherExpiringAccounts -eq $true)
{
	(get-date -Format hh:mm:ss) + (": Running report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Working on Expiring Accounts Report..." -ForegroundColor Yellow
	$Users | Where-Object { $_.AccountExpirationDate -ne $Null } | Foreach-object{
		
		$UserExpiresIn = (New-TimeSpan -Start (get-date) -End $_.AccountExpirationDate).Days
		If (($UserExpiresIn -ge 0) -and ($UserExpiresIn -le $ExpiringAccountDays))
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) will expire in $ExpiringAccountDays days or less") | Out-File ($DirPath + "\" + "Log.txt") -Append
			Write-Host "$($_.Name) will expire in $ExpiringAccountDays days or less"
			$obj = [PSCustomObject]@{
				
				'Name' = "$($_.Name)"
				'AccountExpiringIn' = "$UserExpiresIn Days"
				'LastLogonDate' = "$($_.LastLogonDate)"
				'EmailAddress' = "$($_.emailaddress)"
				'Enabled' = "$($_.Enabled)"
				'SamAccountName' = "$($_.SamAccountName)"
			}
			(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$ExpiringUsersTable.Add($obj)
			(get-date -Format hh:mm:ss) + (": Incrimenting expiring account integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$Int_ExpiringAccounts++
			
		}
	}
	if (($ExpiringUsersTable).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": No accounts were found to expire") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information																    = "Information: No accounts will expire in $($ExpiringAccountDays) days or less"
		}
		$ExpiringUsersTable.Add($obj)
	}
	(get-date -Format hh:mm:ss) + (": Finished with expiring accounts report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Finished Expiring Accounts Report..." -ForegroundColor Cyan
}
Else
{
	(get-date -Format hh:mm:ss) + (": Report set to not run") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Report to Gather Expiring Users is set to False, skipping..." -ForegroundColor Yellow
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$ExpiringUsersTable.Add($obj)
	(get-date -Format hh:mm:ss) + (": Setting integer counter for expiring accounts to N/A") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Int_ExpiringAccounts = "N/A"
}

$LockedOutUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Locked out users") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherLockedOutUsers -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Working on locked out users report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Working on Locked Out Account Report..." -ForegroundColor Yellow
	
	$Users | Where-Object { ($_.Enabled -eq $True) -and ($_.LockedOut -eq $True) } | ForEach-Object{
		(get-date -Format hh:mm:ss) + (": $($_.Name) is currently locked out") | Out-File ($DirPath + "\" + "Log.txt") -Append
		Write-Host "$($_.Name) is currently locked out" -ForegroundColor White
		$obj = [PSCustomObject]@{
			'Name' = "$($_.Name)"
			'LastLogonDate' = "$($_.LastLogonDate)"
			'EmailAddress' = "$($_.emailaddress)"
			'LockedOut' = "$($_.LockedOut)"
			'UPN'  = "$($_.UserPrincipalName)"
			'Enabled' = "$($_.Enabled)"
			'PasswordNeverExpires' = "$($_.PasswordNeverExpires)"
			'SamAccountName' = "$($_.SamAccountName)"
		}
		(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$LockedOutUsersTable.Add($obj)
		(get-date -Format hh:mm:ss) + (": Incrimenting locked out user integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Int_LockedOut++
		
	}
	Write-Host "Finished Locked Out Accounts Report..." -ForegroundColor Cyan
	if (($LockedOutUsersTable).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": No users are currently locked out") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No users are currently locked out'
		}
		$LockedOutUsersTable.Add($obj)
	}
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Report set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Report to Gather Locked Out Users is set to False, skipping..." -ForegroundColor Yellow
	
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$LockedOutUsersTable.Add($obj)
	(get-date -Format hh:mm:ss) + (": Setting locked out users integer counter to N/A") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Int_LockedOut = "N/A"
	
}

$InactiveUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Inactive Users Report") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherInactiveUsers -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Getting all inactive") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Working on Inactive Users Report..." -ForegroundColor Yellow
	
	#If lastlogondate is not empty, and less than or equal to XX days and enabled
	$Users | Where-Object { ($_.Lastlogondate -notlike $null) -and ($_.Enabled -eq $true) } | ForEach-Object{
		
		$LastLogonDate = [datetime]::FromFileTime($_.LastLogonTimeStamp)
		(get-date -Format hh:mm:ss) + (": $($_.Name)'s last logon was $LastLogonDate") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Today = (GET-DATE)
		
		
		(get-date -Format hh:mm:ss) + (": Getting the days since the users last logon") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$DaysSince = ((NEW-TIMESPAN –Start $LastLogonDate –End $Today).Days).ToString()
		(get-date -Format hh:mm:ss) + (": $($_.Name)'s last logon was $DaysSince days ago") | Out-File ($DirPath + "\" + "Log.txt") -Append
		If ($InactiveDays -le $DaysSince)
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) is an inactive user") | Out-File ($DirPath + "\" + "Log.txt") -Append
			Write-Host "$($_.Name) is an inactive user" -ForegroundColor White
			$obj = [PSCustomObject]@{
				
				'Name' = "$($_.Name)"
				'LastLogonDaysAgo' = "$DaysSince"
				'EmailAddress' = "$($_.emailaddress)"
				'LockedOut' = "$($_.LockedOut)"
				'UPN'  = "$($_.UserPrincipalName)"
				'Enabled' = "$($_.Enabled)"
				'PasswordNeverExpires' = "$($_.PasswordNeverExpires)"
				'SamAccountName' = "$($_.SamAccountName)"
			}
			(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$InactiveUsersTable.Add($obj)
			(get-date -Format hh:mm:ss) + (": Incrimenting the inactive user integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$Int_InactiveUsers++
		}
		
		
		
	}
	if (($InactiveUsersTable).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": There are no inactive users") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No inactive accounts'
		}
		$InactiveUsersTable.Add($obj)
	}
	
	Write-Host "Finished with Inactive Users Report..." -ForegroundColor Cyan
}
Else
{
	(get-date -Format hh:mm:ss) + (": Report set to not run") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Report to Gather Inactive Users is set to False, skipping..." -ForegroundColor Yellow
	
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$InactiveUsersTable.Add($obj)
	(get-date -Format hh:mm:ss) + (": Modifying the inactive users integer counter to N/A") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Int_InactiveUsers = "N/A"
	
}


$Table = New-Object 'System.Collections.Generic.List[System.Object]'
$TableExpiresToday = New-Object 'System.Collections.Generic.List[System.Object]'
$TableNoEmailAddress = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Users with passwords expiring") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherPasswordExpiringUsers -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Getting users password expiration report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Running Password Expiration Report...." -ForegroundColor Yellow
	
	(get-date -Format hh:mm:ss) + (": Getting domain password policy") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Getting domain password policy"
	$maxPasswordAge = (Get-ADDefaultDomainPasswordPolicy).MaxPasswordAge
	
	$totalusers = 0
	$Users | Where-Object { ($_.Enabled -ne $false) -and ($_.PasswordNeverExpires -ne $true) -and ($_.PasswordExpired -ne $true) -and ($null -ne $_.PasswordLastSet) } | ForEach-Object{
		$Name = $_.Name
		$totalusers++
		(get-date -Format hh:mm:ss) + (": Working on $($_.Name)") | Out-File ($DirPath + "\" + "Log.txt") -Append
		Write-Host "Working on $Name..." -ForegroundColor White
		
		#Get Password last set date
		(get-date -Format hh:mm:ss) + (": Getting users password last set date") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$passwordSetDate = ($_.PasswordLastSet)
		(get-date -Format hh:mm:ss) + (": Password last set date is $passwordSetDate") | Out-File ($DirPath + "\" + "Log.txt") -Append
		#Check for Fine Grained Passwords
		(get-date -Format hh:mm:ss) + (": Checking for fine grained password policies") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$PasswordPol = (Get-ADUserResultantPasswordPolicy $_.ObjectGUID -ErrorAction silentlycontinue)
		if (($PasswordPol) -ne $null)
		{
			(get-date -Format hh:mm:ss) + (": Fine grained password policies applied to the user, getting new max password age") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$maxPasswordAge = ($PasswordPol).MaxPasswordAge
		}
		
		(get-date -Format hh:mm:ss) + (": Getting password expiration date for user") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$expireson = $passwordsetdate + $maxPasswordAge
		$today = (get-date)
		#Gets the count on how many days until the password expires and stores it in the $daystoexpire var
		$daystoexpire = (New-TimeSpan -Start $today -End $Expireson).Days
		(get-date -Format hh:mm:ss) + (": Password will expire in $daystoexpire days") | Out-File ($DirPath + "\" + "Log.txt") -Append
		(get-date -Format hh:mm:ss) + (": Seeing which query the users password falls under") | Out-File ($DirPath + "\" + "Log.txt") -Append
		
		If (($daystoexpire -gt "0") -and ($daystoexpire -le $expireindays))
		{
			(get-date -Format hh:mm:ss) + (": The password expires in 1 or more days (not negative), and is less than or equal to $expireindays days") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$var = 1
			(get-date -Format hh:mm:ss) + (": Var = $var") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			(get-date -Format hh:mm:ss) + (": User matches query") | Out-File ($DirPath + "\" + "Log.txt") -Append
			(get-date -Format hh:mm:ss) + (": Incrimenting the less than table integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$Int_LessThan++
			$obj = [PSCustomObject]@{
				
				'Name' = "$Name"
				'DaysUntilExpired' = "$daystoexpire"
				'SamAccountName' = "$($_.SamAccountName)"
				'EmailAddress' = "$($_.emailaddress)"
			}
			(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$Table.Add($obj)
			
			
		}
		ElseIf (($daystoexpire -eq "0") -and ($daystoexpire -lt $expireindays))
		{
			(get-date -Format hh:mm:ss) + (": The password expires in 0 days (today), and is less than $expireindays days") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$var = 2
			(get-date -Format hh:mm:ss) + (": Var = $var") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			(get-date -Format hh:mm:ss) + (": User matches query") | Out-File ($DirPath + "\" + "Log.txt") -Append
			(get-date -Format hh:mm:ss) + (": Incrimenting the expires today incriment counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$Int_Today++
			$obj = [PSCustomObject]@{
				
				'Name' = "$Name"
				'DaysUntilExpired' = "$daystoexpire"
				'SamAccountName' = "$($_.SamAccountName)"
				'EmailAddress' = "$($_.emailaddress)"
			}
			(get-date -Format hh:mm:ss) + (": Adding user to expires today table") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$TableExpiresToday.Add($obj)
		}
		Else
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) does not match any query, skipping") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
		}
		
		(get-date -Format hh:mm:ss) + (": Checking to see if sending end user password expiration email") | Out-File ($DirPath + "\" + "Log.txt") -Append
		If (($SendEndUserPasswordExpirationNoticeEmails -eq $true) -and ($daystoexpire -ge "0") -and ($daystoexpire -lt $expireindays))
		{
			(get-date -Format hh:mm:ss) + (": $($_.name) matches query, will be sending email notification") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			(get-date -Format hh:mm:ss) + (": Getting email for $($_.Name)") | Out-File ($DirPath + "\" + "Log.txt") -Append
			Write-Host "Getting e-mail address for $Name..." -ForegroundColor Yellow
			$emailaddress = $_.emailaddress
			If (!($emailaddress))
			{
				(get-date -Format hh:mm:ss) + (": $Name has no E-Mail address listed, looking at their proxyaddresses attribute...") | Out-File ($DirPath + "\" + "Log.txt") -Append
				Write-Host "$Name has no E-Mail address listed, looking at their proxyaddresses attribute..." -ForegroundColor Red
				Try
				{
					$emailaddress = ($_.proxyaddresses | Where-Object { $_ -cmatch '^SMTP' }).Trim("SMTP:")
					(get-date -Format hh:mm:ss) + (": Email set to $emailaddress") | Out-File ($DirPath + "\" + "Log.txt") -Append
				}
				Catch
				{
					$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
				}
			}
			If (!($emailaddress))
			{
				(get-date -Format hh:mm:ss) + (": No Email Address found for user") | Out-File ($DirPath + "\" + "Log.txt") -Append
				Write-Host "$Name has no email addresses to send an e-mail to!" -ForegroundColor Red
				#Don't continue on as we can't email $Null, but if there is an e-mail found it will email that address
				(get-date -Format hh:mm:ss) + (": WARNING: No email found for $Name") | Out-File ($DirPath + "\" + "Log.txt") -Append
				(get-date -Format hh:mm:ss) + (": NO CONTINUING WITH $($_.Name)") | Out-File ($DirPath + "\" + "Log.txt") -Append
				
				(get-date -Format hh:mm:ss) + (": Incrimenting No EMail Incriment counter by 1") | Out-File ($DirPath + "\" + "Log.txt") -Append
				
				$Int_NoEmail++
				$obj = [PSCustomObject]@{
					
					'Name' = "$Name"
					'DaysUntilExpired' = "$daystoexpire"
					'SamAccountName' = "$($_.SamAccountName)"
					'EmailAddress' = "$emailaddress"
				}
				(get-date -Format hh:mm:ss) + (": Adding user to no email address table") | Out-File ($DirPath + "\" + "Log.txt") -Append
				$TableNoEmailAddress.Add($obj)
				
			}
			elseif (($null -ne $emailaddress) -and ($var -eq 1))
			{
				(get-date -Format hh:mm:ss) + (": $($_.Name) has an email and password is expiring in greater than 0 days") | Out-File ($DirPath + "\" + "Log.txt") -Append
				
				(get-date -Format hh:mm:ss) + (": Sending expiry notice email to $Name") | Out-File ($DirPath + "\" + "Log.txt") -Append
				Write-Host "Sending Password expiry email to $name" -ForegroundColor Yellow
				
				$SmtpClient = new-object system.net.mail.smtpClient
				$MailMessage = New-Object system.net.mail.mailmessage
				
				#Who is the e-mail sent from
				$mailmessage.From = $FromEmail
				#SMTP server to send email
				$SmtpClient.Host = $SMTPHost
				#SMTP SSL
				$SMTPClient.EnableSsl = $true
				#SMTP credentials
				$SMTPClient.Credentials = $cred
				#Send e-mail to the users email
				(get-date -Format hh:mm:ss) + (": Sending email to $emailaddress") | Out-File ($DirPath + "\" + "Log.txt") -Append
				$mailmessage.To.add("$emailaddress")
				#Email subject
				$mailmessage.Subject = "Your password will expire in $daystoexpire days"
				#Notification email on delivery / failure
				$MailMessage.DeliveryNotificationOptions = ("onSuccess", "onFailure")
				$MailMessage.IsBodyHtml = $True
				$mailmessage.Body =
				"<pr>Good Morning $Name,
<br>
<br>
Your Domain password will expire in $daystoexpire days. Please change it as soon as possible.
<br>
<br>
To change your password, follow the method below:
<br>
</p>
<p>
<b> On your Windows computer</b>
<ol>
  <li>If you are not in the office, logon and connect to VPN.</li>
		<ul>
  			<li><b>NOTE:</b> Users in the main office, login as normal. If you are working remotely, you must connect to the VPN before changing your password.</li>
		</ul>
  <li>Log onto your computer as usual and make sure you are connected to the internet.</li>
  <li>Press Ctrl-Alt-Del and click on ""Change Password"".</li>
  <li>Fill in your old password and set a new password.  See the password requirements below.</li>
  <li>Press OK to return to your desktop. .</li>
</ol>
</p>
<p>
<br> 
<b>The new password must meet the minimum requirements set forth in our corporate policies including:</b>
<ol>
  <li>It must be at least 8 characters long.</li>
  <li>It must contain at least one character from 3 of the 4 following groups of characters:</li>
		<ul>
  			<li>Uppercase letters (A-Z)</li>
  			<li>Lowercase letters (a-z)</li>
  			<li>Numbers (0-9)</li>
			<li>Symbols (!@#`$%^&*...)</li>
		</ul>
  <li>It cannot match any of your past 10 passwords.</li>
  <li>It cannot contain characters which match 3 or more consecutive characters of your username.</li>
  <li>If you attempt to change a PW more than once in 24 hours, you well get a password change error message. 
		<ul>
  			<li>Please wait the full day between password changes or contact the IT Service desk if you need it changed immediately with the contact info below.</li>
		</ul>
	</li>
</ol>
</p>
<p>
<br>
If you used the 'remember my password' feature with some of our web applications (e.g. Connect), you will need to ensure that you change that password as well if you want the system to remember your new password.
<br> 
<br> 
If you have any questions please contact the Support team at <a href='mailto:Helpdesk@TheLazyAdministrator.com?Subject=Account Password Expiration' >Helpdesk@TheLazyAdministrator.com</a> or call us at 800-872-9622 ext. 2222
<br>
<br> 
Thanks and have a great $((get-date).dayofweek)!
<br> 
IT Department
<br> 
<a href='mailto:Helpdesk@TheLazyAdministrator.com?Subject=Account Password Expiration' >Helpdesk@TheLazyAdministrator.com</a>
<br> 
800-872-9622 ext. 2222
<br> 
<br>
<img src='https://www.thelazyadministrator.com/wp-content/uploads/2018/08/ewwewewe.png' alt='The Lazy Administrator'>
</pr>"
				Write-Host "Sending E-mail to $emailaddress..." -ForegroundColor Green
				Try
				{
					$smtpclient.Send($mailmessage)
				}
				Catch
				{
					$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
				}
			}
			elseif (($null -ne $emailaddress) -and ($var -eq 2))
			{
				(get-date -Format hh:mm:ss) + (": $($_.Name) has an email and password is expiring today") | Out-File ($DirPath + "\" + "Log.txt") -Append
				
				
				
				(get-date -Format hh:mm:ss) + (": Sending expiry notice email to $Name") | Out-File ($DirPath + "\" + "Log.txt") -Append
				Write-Host "Sending Password expiry email to $name" -ForegroundColor Yellow
				
				$SmtpClient = New-Object system.net.mail.smtpClient
				$MailMessage = New-Object system.net.mail.mailmessage
				
				#Who is the e-mail sent from
				$mailmessage.From = $FromEmail
				#SMTP server to send email
				$SmtpClient.Host = $SMTPHost
				#SMTP SSL
				$SMTPClient.EnableSsl = $true
				#SMTP credentials
				$SMTPClient.Credentials = $cred
				#Send e-mail to the users email
				(get-date -Format hh:mm:ss) + (": Sending email to $emailaddress") | Out-File ($DirPath + "\" + "Log.txt") -Append
				$mailmessage.To.add("$emailaddress")
				#Email subject
				$mailmessage.Subject = "Your password will expire TODAY!"
				#Notification email on delivery / failure
				$MailMessage.DeliveryNotificationOptions = ("onSuccess", "onFailure")
				#Send e-mail with high priority
				$MailMessage.Priority = "High"
				$MailMessage.IsBodyHtml = $True
				"<pr>Good Morning $Name,
<br>
<br>
Your Domain password will expire <b>TODAY</b>! You must change it as soon as possible.
<br>
<br>
To change your password, follow the method below:
<br>
</p>
<p>
<b> On your Windows computer</b>
<ol>
  <li>If you are not in the office, logon and connect to VPN.</li>
		<ul>
  			<li><b>NOTE:</b> Users in the main office, login as normal. If you are working remotely, you must connect to the VPN before changing your password.</li>
		</ul>
  <li>Log onto your computer as usual and make sure you are connected to the internet.</li>
  <li>Press Ctrl-Alt-Del and click on ""Change Password"".</li>
  <li>Fill in your old password and set a new password.  See the password requirements below.</li>
  <li>Press OK to return to your desktop. .</li>
</ol>
</p>
<p>
<br> 
<b>The new password must meet the minimum requirements set forth in our corporate policies including:</b>
<ol>
  <li>It must be at least 8 characters long.</li>
  <li>It must contain at least one character from 3 of the 4 following groups of characters:</li>
		<ul>
  			<li>Uppercase letters (A-Z)</li>
  			<li>Lowercase letters (a-z)</li>
  			<li>Numbers (0-9)</li>
			<li>Symbols (!@#`$%^&*...)</li>
		</ul>
  <li>It cannot match any of your past 10 passwords.</li>
  <li>It cannot contain characters which match 3 or more consecutive characters of your username.</li>
  <li>If you attempt to change a PW more than once in 24 hours, you well get a password change error message. 
		<ul>
  			<li>Please wait the full day between password changes or contact the IT Service desk if you need it changed immediately with the contact info below.</li>
		</ul>
	</li>
</ol>
</p>
<p>
<br>
If you used the 'remember my password' feature with some of our web applications (e.g. Connect), you will need to ensure that you change that password as well if you want the system to remember your new password.
<br> 
<br> 
If you have any questions please contact the Support team at <a href='mailto:Helpdesk@TheLazyAdministrator.com?Subject=Account Password Expiration' >Helpdesk@TheLazyAdministrator.com</a> or call us at 800-872-9622 ext. 2222
<br>
<br> 
Thanks and have a great $((get-date).dayofweek)!
<br> 
IT Department
<br> 
<a href='mailto:Helpdesk@TheLazyAdministrator.com?Subject=Account Password Expiration' >Helpdesk@TheLazyAdministrator.com</a>
<br> 
800-872-9622 ext. 2222
<br> 
<br>
<img src='https://www.thelazyadministrator.com/wp-content/uploads/2018/08/ewwewewe.png' alt='The Lazy Administrator'>
</pr>"
				Write-Host "Sending E-mail to $emailaddress..." -ForegroundColor Green
				Try
				{
					$smtpclient.Send($mailmessage)
				}
				Catch
				{
					$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
				}
			}
			Else
			{
				(get-date -Format hh:mm:ss) + (": WARNING! Something went wrong with this user. Please check query. Name: $($_.Name); EMAIL: $emailaddress") | Out-File ($DirPath + "\" + "Log.txt") -Append
				
			}
		}
		Else
		{
			(get-date -Format hh:mm:ss) + (": End user password expiration email set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
		}
	}
	
	if (($Table).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": No users were found with passwords expiring soon") | Out-File ($DirPath + "\" + "Log.txt") -Append
		
		$Obj = [PSCustomObject]@{
			
			Information	     = "Information: No Users were found to have expiring passwords in $expireindays days"
		}
		$Table.Add($obj)
	}
	if (($TableExpiresToday).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": WARNING: No users were found with passwords expiring today") | Out-File ($DirPath + "\" + "Log.txt") -Append
		
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No Users have passwords expiring today'
		}
		$TableExpiresToday.Add($obj)
	}
	if (($TableNoEmailAddress).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": WARNING: No users were found with passwords expiring and no valid email address") | Out-File ($DirPath + "\" + "Log.txt") -Append
		
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No Users have passwords expiring and have no email address or End-User Email Notifications was set to False'
		}
		$TableNoEmailAddress.Add($obj)
	}
	
}
Else
{
	
	Write-Host "Report to Gather Password Expiring Users is set to False, skipping..." -ForegroundColor Yellow
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$Table.Add($obj)
	$Int_LessThan = "N/A"
	
	
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$TableExpiresToday.Add($obj)
	$Int_Today = "N/A"
	
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$TableNoEmailAddress.Add($obj)
	$Int_NoEmail = "N/A"
	
	
}

$PasswordNeverExpiresTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Password never expires") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherUsersWithPasswordsNeverSetToExpire -eq $true)
{
	(get-date -Format hh:mm:ss) + (": Getting all users with passwords set to never expire") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Working on Password Never Expires Report..." -ForegroundColor Yellow
	$Users | Where-Object { $_.PasswordNeverExpires -eq $true } | Foreach-object{
		(get-date -Format hh:mm:ss) + (": $($_.Name)'s password is currently set to never expire") | Out-File ($DirPath + "\" + "Log.txt") -Append
		Write-Host "$($_.Name)'s password is currently set to never expire"
		$obj = [PSCustomObject]@{
			'Name' = "$($_.Name)"
			'LastLogonDate' = "$($_.LastLogonDate)"
			'EmailAddress' = "$($_.emailaddress)"
			'LockedOut' = "$($_.LockedOut)"
			'Enabled' = "$($_.Enabled)"
			'PasswordNeverExpires' = "$($_.PasswordNeverExpires)"
			'PasswordLastSet' = "$($_.PasswordLastSet)"
			'SamAccountName' = "$($_.SamAccountName)"
		}
		(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$PasswordNeverExpiresTable.Add($obj)
		(get-date -Format hh:mm:ss) + (": Incrimenting the PW Never Expire integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Int_PWNeverExpires++
	}
	if (($PasswordNeverExpiresTable).Count -eq 0)
	{
		(get-date -Format hh:mm:ss) + (": No users were found with passwords set to never expire") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No new users have their passwords set to never expire'
		}
		$PasswordNeverExpiresTable.Add($obj)
	}
}
Else
{
	(get-date -Format hh:mm:ss) + (": Report set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Report to Users with Passwords set to Never Expire is set to False, skipping..." -ForegroundColor Yellow
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$PasswordNeverExpiresTable.Add($obj)
	(get-date -Format hh:mm:ss) + (": Modifying the PW Never expire integer counter to N/A") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Int_PWNeverExpires = "N/A"
	
}

$NewCreatedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Newly created users") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherNewlyCreatedUsers -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Gathering newly created users") | Out-File ($DirPath + "\" + "Log.txt") -Append
	#Get newly created users
	$When = ((Get-Date).AddDays(- $UserCreatedDays)).Date
	$Users | Where-Object { $_.whenCreated -ge $When } | ForEach-Object {
		(get-date -Format hh:mm:ss) + (": $($_.Name) is a newly created user") | Out-File ($DirPath + "\" + "Log.txt") -Append
		
		$obj = [PSCustomObject]@{
			
			'Name' = $_.Name
			'Enabled' = $_.Enabled
			'CreationDate' = $_.whenCreated
		}
		(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$NewCreatedUsersTable.Add($obj)
		(get-date -Format hh:mm:ss) + (": Incrimenting the New Users table integer counter") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Int_NewAccounts++
	}
	if (($NewCreatedUsersTable).Count -lt 1)
	{
		(get-date -Format hh:mm:ss) + (": No new users have been created recently") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No new users have been recently created'
		}
		$NewCreatedUsersTable.Add($obj)
	}
}
Else
{
	(get-date -Format hh:mm:ss) + (": Report set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append
	Write-Host "Report to Gather Newly Created Users is set to False, skipping..." -ForegroundColor Yellow
	
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$NewCreatedUsersTable.Add($obj)
	(get-date -Format hh:mm:ss) + (": Modifying the new users integer counter to N/A") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Int_NewAccounts = "N/A"
}

$UsersNoManagersTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": REPORT: Users Without Managers") | Out-File ($DirPath + "\" + "Log.txt") -Append
If ($GatherUsersWithoutManagers -eq $true)
{
	(get-date -Format hh:mm:ss) + (": Gathering users without managers") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$users | Where-object { $null -eq $_.manager } | ForEach-Object{
		(get-date -Format hh:mm:ss) + (": $($_.Name) is a user without a manager") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Int_NoManager++
		$obj = [PSCustomObject]@{
			
			'Name' = $_.Name
			'PasswordLastSet' = $_.PasswordLastSet
			'LastLogon' = $_.LastLogonDate
			'Enabled' = $_.Enabled
		}
		(get-date -Format hh:mm:ss) + (": Adding user to table") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$UsersNoManagersTable.Add($obj)
	}
	if (($UsersNoManagersTable).Count -lt 1)
	{
		(get-date -Format hh:mm:ss) + (": No users were found to have no managers") | Out-File ($DirPath + "\" + "Log.txt") -Append
		$Obj = [PSCustomObject]@{
			
			Information = 'Information: No users were found to have no managers'
		}
		$UsersNoManagersTable.Add($obj)
	}
}
Else
{
	$Int_NoManager = "N/A"
	(get-date -Format hh:mm:ss) + (": Report set to false, not running") | Out-File ($DirPath + "\" + "Log.txt") -Append
	$Obj = [PSCustomObject]@{
		
		Information = 'Information: No Data; Report did not run'
	}
	$UsersNoManagersTable.Add($obj)
}

$EnabledDisabledUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$ProtectedUsersTable = New-Object 'System.Collections.Generic.List[System.Object]'
$PasswordExpirationTable = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": WARNING: Checking to see if disable graphs is set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append

If ($DisableGraphs -eq $false)
{
	(get-date -Format hh:mm:ss) + (": WARNING: Disable graphs is set to false") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Graphs enabled, gathering data..."
	
	
	##--Enabled users vs Disabled Users PIE CHART--##
	#Basic Properties 
	$EnabledDisabledUsersPieObject = Get-HTMLPieChartObject
	$EnabledDisabledUsersPieObject.Title = "Enabled vs Disabled Users"
	$EnabledDisabledUsersPieObject.Size.Height = 250
	$EnabledDisabledUsersPieObject.Size.width = 250
	$EnabledDisabledUsersPieObject.ChartStyle.ChartType = 'doughnut'
	
	#you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$EnabledDisabledUsersPieObject.ChartStyle.ColorSchemeName = 'Random'
	
	#Data defintion you can reference any column from name and value from the  dataset.  
	#Name and Count are the default to work with the Group function.
	$EnabledDisabledUsersPieObject.DataDefinition.DataNameColumnName = 'Name'
	$EnabledDisabledUsersPieObject.DataDefinition.DataValueColumnName = 'Count'
	
	
	##--PasswordNeverExpires PIE CHART--##
	#Basic Properties 
	$PWExpiresUsersTable = Get-HTMLPieChartObject
	$PWExpiresUsersTable.Title = "Password Never Expires Status"
	$PWExpiresUsersTable.Size.Height = 250
	$PWExpiresUsersTable.Size.Width = 250
	$PWExpiresUsersTable.ChartStyle.ChartType = 'doughnut'
	
	
	#you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PWExpiresUsersTable.ChartStyle.ColorSchemeName = 'Random'
	
	#Data defintion you can reference any column from name and value from the  dataset.  
	#Name and Count are the default to work with the Group function.
	$PWExpiresUsersTable.DataDefinition.DataNameColumnName = 'Name'
	$PWExpiresUsersTable.DataDefinition.DataValueColumnName = 'Count'
	
	
	##--USERS Protection PIE CHART--##
	#Basic Properties 
	$PieObjectProtectedUsers = Get-HTMLPieChartObject
	$PieObjectProtectedUsers.Title = "Users Protected from Deletion"
	$PieObjectProtectedUsers.Size.Height = 250
	$PieObjectProtectedUsers.Size.width = 250
	$PieObjectProtectedUsers.ChartStyle.ChartType = 'doughnut'
	
	#you can also ask for a random scheme.  Which also happens ifyou have too many records for the scheme
	$PieObjectProtectedUsers.ChartStyle.ColorSchemeName = 'Random'
	
	#Data defintion you can reference any column from name and value from the  dataset.  
	#Name and Count are the default to work with the Group function.
	$PieObjectProtectedUsers.DataDefinition.DataNameColumnName = 'Name'
	$PieObjectProtectedUsers.DataDefinition.DataValueColumnName = 'Count'
	
	$UserEnabled = 0
	$UserDisabled = 0
	
	(get-date -Format hh:mm:ss) + (": Gathering data for enabled/disabled users") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	$Users | ForEach-Object{
		If ($_.Enabled -eq $False)
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) is disabled") | Out-File ($DirPath + "\" + "Log.txt") -Append
			$UserDisabled++
		}
		Else
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) is enabled") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$UserEnabled++
		}
		
	}
	
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Enabled'
		'Count' = $UserEnabled
	}
	
	$EnabledDisabledUsersTable.Add($objULic)
	
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Disabled'
		'Count' = $UserDisabled
	}
	
	$EnabledDisabledUsersTable.Add($objULic)
	
	#Protected PieChart
	#Data for protected users pie graph
	$ProtectedUsers = 0
	$NonProtectedUsers = 0
	(get-date -Format hh:mm:ss) + (": Gathering data for users protected against accidental deletion") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	$Users | ForEach-Object {
		If ($_.ProtectedFromAccidentalDeletion -eq $true)
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) is protected") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$ProtectedUsers++
		}
		Else
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name) is not protected") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$NonProtectedUsers++
		}
	}
	
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Protected'
		'Count' = $ProtectedUsers
	}
	$ProtectedUsersTable.Add($objULic)
	
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Not Protected'
		'Count' = $NonProtectedUsers
	}
	$ProtectedUsersTable.Add($objULic)
	
	
	#PasswordNeverExpires
	#Data for users password expires pie graph
	$UserPasswordExpires = 0
	$UserPasswordNeverExpires = 0
	(get-date -Format hh:mm:ss) + (": Gathering data for users with passwords set to never expire") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	$Users | ForEach-Object {
		If ($_.PasswordNeverExpires -eq $False)
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name)'s password expires") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$UserPasswordExpires++
		}
		Else
		{
			(get-date -Format hh:mm:ss) + (": $($_.Name)'s password never expires") | Out-File ($DirPath + "\" + "Log.txt") -Append
			
			$UserPasswordNeverExpires++
		}
		
	}
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Password Expires'
		'Count' = $UserPasswordExpires
	}
	$PasswordExpirationTable.Add($objULic)
	
	$objULic = [PSCustomObject]@{
		
		'Name'  = 'Password Never Expires'
		'Count' = $UserPasswordNeverExpires
	}
	
	$PasswordExpirationTable.Add($objULic)
	
	
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Disable graphs is set to true") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
}


$FinalReport = New-Object 'System.Collections.Generic.List[System.Object]'
(get-date -Format hh:mm:ss) + (": Checking to see if summary email is requested") | Out-File ($DirPath + "\" + "Log.txt") -Append

If ($SendSummaryEmailNotification -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Summary email is requested, generating report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Generating HTML Report"
	$tabarray = @('Users')
	$FinalReport.Add($(Get-HTMLOpenPage -TitleText $ReportTitle -LeftLogoString $CompanyLogo -RightLogoString $RightLogo))
	$FinalReport.Add($(Get-HTMLTabHeader -TabNames $tabarray))
	$FinalReport.Add($(Get-HTMLTabContentopen -TabName $tabarray[0] -TabHeading ("Report: " + (Get-Date -Format MM-dd-yyyy))))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users With Passwords Expiring Today"))
	$FinalReport.Add($(Get-HTMLContentDataTable $TableExpiresToday))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users With Passwords Expiring in $expireindays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $Table))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users With Passwords Expiring But No Email Address"))
	$FinalReport.Add($(Get-HTMLContentDataTable $TableNoEmailAddress))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users With Passwords Set to Never Expire"))
	$FinalReport.Add($(Get-HTMLContentDataTable $PasswordNeverExpiresTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Inactive Users"))
	$FinalReport.Add($(Get-HTMLContentDataTable $InactiveUsersTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Locked Out Users"))
	$FinalReport.Add($(Get-HTMLContentDataTable $LockedOutUsersTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users With No Manager Set"))
	$FinalReport.Add($(Get-HTMLContentDataTable $UsersNoManagersTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	
	$FinalReport.Add($(Get-HTMLContentOpen -BackgroundShade 1 -HeaderText "Accounts Created in $UserCreatedDays Days or Less"))
	$FinalReport.Add($(Get-HTMLContentDataTable $NewCreatedUsersTable -HideFooter))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	
	$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Expiring User Accounts"))
	$FinalReport.Add($(Get-HTMLContentDataTable $ExpiringUsersTable))
	$FinalReport.Add($(Get-HTMLContentClose))
	
	If ($DisableGraphs -eq $false)
	{
		$FinalReport.Add($(Get-HTMLContentOpen -HeaderText "Users Charts"))
		$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 1 -ColumnCount 3))
		$FinalReport.Add($(Get-HTMLPieChart -ChartObject $EnabledDisabledUsersPieObject -DataSet $EnabledDisabledUsersTable))
		$FinalReport.Add($(Get-HTMLColumnClose))
		$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 2 -ColumnCount 3))
		$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PWExpiresUsersTable -DataSet $PasswordExpirationTable))
		$FinalReport.Add($(Get-HTMLColumnClose))
		$FinalReport.Add($(Get-HTMLColumnOpen -ColumnNumber 3 -ColumnCount 3))
		$FinalReport.Add($(Get-HTMLPieChart -ChartObject $PieObjectProtectedUsers -DataSet $ProtectedUsersTable))
		$FinalReport.Add($(Get-HTMLColumnClose))
		$FinalReport.Add($(Get-HTMLContentClose))
		$FinalReport.Add($(Get-HTMLTabContentClose))
		
	}
	
	$FinalReport.Add($(Get-HTMLTabContentClose))
	$FinalReport.Add($(Get-HTMLClosePage))
	
	Save-HTMLReport -ReportContent $FinalReport -ReportName "IT_Report" -ReportPath $DirPath
	(get-date -Format hh:mm:ss) + (": Finished HTML Report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Finished HTML Report"
	(get-date -Format hh:mm:ss) + (": Creating summary report") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Creating summary report"
	$SmtpClient = new-object system.net.mail.smtpClient
	$MailMessage = New-Object system.net.mail.mailmessage
	
	$file = "$DirPath\IT_Report.html"
	(get-date -Format hh:mm:ss) + (": HTML report will be at $file") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	$att = new-object Net.Mail.Attachment($file)
	
	#Who is the e-mail sent from
	$mailmessage.From = $FromEmail
	#SMTP server to send email
	$SmtpClient.Host = $SMTPHost
	#SMTP SSL
	$SMTPClient.EnableSsl = $true
	#SMTP credentials
	$SMTPClient.Credentials = $cred
	#Send e-mail to the users email
	$mailmessage.To.add("$SummaryEmailAddress")
	(get-date -Format hh:mm:ss) + (": Sending summary email to $SummaryEmailAddress") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	#Email subject
	$mailmessage.Subject = "Automation Report for $(get-date -Format "dddd, MMMM dd yyyy")"
	#Notification email on delivery / failure
	$MailMessage.DeliveryNotificationOptions = ("onSuccess", "onFailure")
	$MailMessage.IsBodyHtml = $True
	
	
	#Make the HTML tables for users with passwords expiring
	If ($Table -like "*Information: No Users*")
	{
		$daysuntilq = "<p><font color=`"Black`"><b>Users With Passwords Expiring in $expireindays Days or Less:</b></font><br>
There are no users with passwords expiring in $expireindays or less</P>"
	}
	ElseIf ($Table -like "*Information: No Data*")
	{
		$daysuntilq = "<p><font color=`"Black`"><b>Users With Passwords Expiring Today:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($Table.count -gt 0)
	{
		[string]$daysuntilq = [PSCustomObject]$table | Select-Object -Property "Name", "DaysUntilExpired", "SamAccountName" | Sort-Object -Property "DaysUntilExpired" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users With Passwords Expiring in 7 Days or Less:</b></font>"
	}
	Else
	{
		$daysuntilq = "<p><font color=`"Black`"><b>Users With Passwords Expiring in $expireindays Days or Less:</b></font><br>
There are no users with passwords expiring in $expireindays or less</P>"
	}
	
	
	If ($TableExpiresToday -like "*Information: No Users*")
	{
		$daysalreadyq = "<p><font color=`"Black`"><b>Users With Passwords Expiring Today:</b></font><br>
There are no users with passwords expiring today</P>"
	}
	ElseIf ($TableExpiresToday -like "*Information: No Data*")
	{
		$daysalreadyq = "<p><font color=`"Black`"><b>Users With Passwords Expiring Today:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($TableExpiresToday.count -gt 0)
	{
		[string]$daysalreadyq = [PSCustomObject]$TableExpiresToday | Select-Object -Property "Name", "DaysUntilExpired", "SamAccountName" | Sort-Object -Property "DaysUntilExpired" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users With Passwords Expiring Today:</b></font>"
	}
	Else
	{
		$daysalreadyq = "<p><font color=`"Black`"><b>Users With Passwords Expiring Today:</b></font><br>
There are no users with passwords expiring today</P>"
	}
	
	If ($ExpiringUsersTable -like "*Information: No accounts*")
	{
		$HTMLTable_ExpiringUsers = "<p><font color=`"Black`"><b>Users Expiring in $ExpiringAccountDays Days or Less:</b></font><br>
There are no users that will expire in $ExpiringAccountDays or less</P>"
	}
	ElseIf ($ExpiringUsersTable -like "*Information: No Data*")
	{
		$HTMLTable_ExpiringUsers = "<p><font color=`"Black`"><b>Users Expiring in $ExpiringAccountDays Days or Less:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($ExpiringUsersTable.Count -gt 0)
	{
		[string]$HTMLTable_ExpiringUsers = [PSCustomObject]$ExpiringUsersTable | Select-Object -Property "Name", "AccountExpiringIn", "LastLogonDate", "SamAccountName" | Sort-Object -Property "AccountExpiringIn" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users Expiring in $ExpiringAccountDays Days or Less:</b></font>"
	}
	Else
	{
		$HTMLTable_ExpiringUsers = = "<p><font color=`"Black`"><b>Users Expiring in $ExpiringAccountDays Days or Less:</b></font><br>
There are no users that will expire in $ExpiringAccountDays or less</P>"
	}
	
	
	If ($LockedOutUsersTable -like "*Information: No Users*")
	{
		$HTMLTable_LockedOutUsers = "<p><font color=`"Black`"><b>Locked Out Users:</b></font><br>
There are currently no locked out accounts</P>"
	}
	ElseIf ($LockedOutUsersTable -like "*Information: No Data*")
	{
		$HTMLTable_LockedOutUsers = "<p><font color=`"Black`"><b>Locked Out Users:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($LockedOutUsersTable.Count -gt 0)
	{
		[string]$HTMLTable_LockedOutUsers = [PSCustomObject]$LockedOutUsersTable | Select-Object -Property "Name", "LastLogonDate", "LockedOut", "EmailAddress", "SamAccountName" | Sort-Object -Property "Name" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Locked Out Users:</b></font>"
	}
	Else
	{
		$HTMLTable_LockedOutUsers = "<p><font color=`"Black`"><b>Locked Out Users:</b></font><br>
There are currently no locked out accounts</P>"
	}
	
	
	If ($InactiveUsersTable -like "*Information: No Users*")
	{
		$HTMLTable_InactiveUsers = "<p><font color=`"Black`"><b>Inactive Users:</b></font><br>
There are currently no inactive accounts</P>"
	}
	ElseIf ($InactiveUsersTable -like "*Information: No Data*")
	{
		$HTMLTable_InactiveUsers = "<p><font color=`"Black`"><b>Inactive Users:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($InactiveUsersTable.Count -gt 0)
	{
		[string]$HTMLTable_InactiveUsers = [PSCustomObject]$InactiveUsersTable | Select-Object -Property "Name", "LastLogonDaysAgo", "EmailAddress", "SamAccountName" | Sort-Object -Property "LastLogonDaysAgo" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Inactive Users:</b></font>"
	}
	Else
	{
		$HTMLTable_InactiveUsers = "<p><font color=`"Black`"><b>Inactive Users:</b></font><br>
There are currently no inactive accounts</P>"
	}
	
	
	If ($PasswordNeverExpiresTable -like "*Information: No Users*")
	{
		$HTMLTable_PasswordNeverExpires = "<p><font color=`"Black`"><b>Users With Passwords Set to Never Expire:</b></font><br>
There are currently no users with passwords set to never expire</P>"
	}
	ElseIf ($PasswordNeverExpiresTable -like "*Information: No Data*")
	{
		$HTMLTable_PasswordNeverExpires = "<p><font color=`"Black`"><b>Users With Passwords Set to Never Expire:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($PasswordNeverExpiresTable.Count -gt 0)
	{
		[string]$HTMLTable_PasswordNeverExpires = [PSCustomObject]$PasswordNeverExpiresTable | Select-Object -Property "Name", "PasswordLastSet", "LastLogonDate", "EmailAddress", "SamAccountName" | Sort-Object -Property "Name" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users With Passwords Set to Never Expire:</b></font>"
		
	}
	Else
	{
		$HTMLTable_PasswordNeverExpires = "<p><font color=`"Black`"><b>Users With Passwords Set to Never Expire:</b></font><br>
There are currently no users with passwords set to never expire</P>"
	}
	
	
	If ($TableNoEmailAddress -like "*Information: No Users*")
	{
		$HTMLTable_PasswordExpiringNoEmail = "<p><font color=`"Black`"><b>Users With Passwords Expiring in $expireindays Days or Less but No Valid Email Address:</b></font><br>
There are currently no users with passwords expiring in $expireindays days or less with not valid email address</P>"
	}
	ElseIf ($TableNoEmailAddress -like "*Information: No Data*")
	{
		$HTMLTable_PasswordExpiringNoEmail = "<p><font color=`"Black`"><b>Users With Passwords Expiring in $expireindays Days or Less but No Valid Email Address:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($TableNoEmailAddress.Count -gt 0)
	{
		[string]$HTMLTable_PasswordExpiringNoEmail = [PSCustomObject]$PasswordNeverExpiresTable | Select-Object -Property "Name", "DaysUntilExpired", "LastLogonDate", "SamAccountName" | Sort-Object -Property "Name" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users With Passwords Set to Never Expire:</b></font>"
		
	}
	Else
	{
		$HTMLTable_PasswordExpiringNoEmail = "<p><font color=`"Black`"><b>Users With Passwords Expiring in $expireindays Days or Less but No Valid Email Address:</b></font><br>
There are currently no users with passwords expiring in $expireindays days or less with not valid email address</P>"
	}
	
	
	
	If ($NewCreatedUsersTable -like "*Information: No New Users*")
	{
		$HTMLTable_NewCreatedUsersTable = "<p><font color=`"Black`"><b>Accounts Created in $UserCreatedDays Days or Less:</b></font><br>
There were no users created in the last $UserCreatedDays Days</P>"
	}
	ElseIf ($NewCreatedUsersTable -like "*Information: No Data*")
	{
		$HTMLTable_NewCreatedUsersTable = "<p><font color=`"Black`"><b>Accounts Created in $UserCreatedDays Days or Less:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($NewCreatedUsersTable.Count -gt 0)
	{
		[string]$HTMLTable_NewCreatedUsersTable = [PSCustomObject]$NewCreatedUsersTable | Select-Object -Property "Name", "Enabled", "CreationDate" | Sort-Object -Property "Name" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Accounts Created in $UserCreatedDays Days or Less:</b></font>"
		
	}
	Else
	{
		$HTMLTable_NewCreatedUsersTable = "<p><font color=`"Black`"><b>Accounts Created in $UserCreatedDays Days or Less:</b></font><br>
There were no users created in the last $UserCreatedDays Days</P>"
	}
	
	
	
	If ($UsersNoManagersTable -like "*Information: No Users*")
	{
		$HTMLTable_NoManagerUsersTable = "<p><font color=`"Black`"><b>Users Without a Manager Set:</b></font><br>
There are no users without a manager set</P>"
	}
	ElseIf ($UsersNoManagersTable -like "*Information: No Data*")
	{
		$HTMLTable_NoManagerUsersTable = "<p><font color=`"Black`"><b>Users Without a Manager Set:</b></font><br>
No Data; Report did not run</P>"
	}
	ElseIf ($UsersNoManagersTable.Count -gt 0)
	{
		[string]$HTMLTable_NoManagerUsersTable = [PSCustomObject]$UsersNoManagersTable | Select-Object -Property "Name", "LastLogon", "PasswordLastSet", "Enabled" | Sort-Object -Property "Name" | ConvertTo-HTML -head $Head -Body "<font color=`"Black`"><b>Users Without a Manager Set:</b></font>"
		
	}
	Else
	{
		$HTMLTable_NoManagerUsersTable = "<p><font color=`"Black`"><b>Users Without a Manager Set:</b></font><br>
There are no users without a manager set</P>"
	}
	
	#Making vocab correct
	If ($Table -like "*Information: No Data*")
	{
		$MessageTable = "<i>There is no data as the report was set to False.</i></p>"
		
	}
	ElseIf ($Table -like "*Information: No Users*")
	{
		$MessageTable = "<i>There are no users that have password expiring in $expireindays days or less.</i></p>"
		
	}
	ElseIf (($Table.Count -eq 1) -and ($TableExpiresToday -like "*Information: No Users*"))
	{
		$MessageTable = "<i>There is $($Table.count) user that has their password expiring in $expireindays days or less. The password does not expire today.</i></p>"
	}
	ElseIf (($Table.count -eq 1) -and ($TableExpiresToday.count -eq 1))
	{
		$MessageTable = "<i>There is 1 user that has their password expiring in $expireindays days or less and that password expires today!</i></p>"
	}
	ElseIf (($Table.count -gt 1) -and ($TableExpiresToday -like "*Information: No Users*"))
	{
		$MessageTable = "<i>There are $($Table.Count) users that have their passwords expiring in $expireindays days or less. Out of those $($Table.count), there are 0 users that have passwords expiring today.</i></p>"
	}
	ElseIf (($Table.count -gt 1) -and ($TableExpiresToday.count -eq 1))
	{
		$MessageTable = "<i>There are $($Table.count) users that have password expiring in $expireindays days or less. Out of those $($Table.count), there are $($TableExpiresToday.count) user that has their passwords expiring today.</i></p>"
	}
	ElseIf (($Table.count -gt 1) -and ($TableExpiresToday -gt 1))
	{
		$MessageTable = "<i>There are $($Table.count) users that have password expiring in $expireindays days or less. Out of those $($Table.count), there are $($TableExpiresToday.count) user that has their passwords expiring today.</i></p>"
	}
	Else
	{
		$MessageTable = "<i>There are $($Table.count) users that have password expiring in $expireindays days or less. Out of those $($Table.count), there are $($TableExpiresToday.count) user that has their passwords expiring today.</i></p>"
	}
	
	
	
	If ($LockedOutUsersTable -like "*Information: No Users*")
	{
		$MessagelockedOut = "<i>There are no users that are currently locked out.</i></p>"
		
	}
	ElseIf ($LockedOutUsersTable -like "*Information: No Data*")
	{
		$MessagelockedOut = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($LockedOutUsersTable.Count) -eq 1)
	{
		$MessagelockedOut = "<i>There is $($LockedOutUsersTable.count) user that is currently locked out.</i></p>"
	}
	Else
	{
		$MessagelockedOut = "<i>There are $($LockedOutUsersTable.count) users that are currently locked out.</i></p>"
	}
	
	
	If ($ExpiringUsersTable -like "*Information: No Accounts*")
	{
		$MessageExpiringAccounts = "<i>There are no users that are expiring in $ExpiringAccountDays days or less.</i></p>"
		
	}
	ElseIf ($ExpiringUsersTable -like "*Information: No Data*")
	{
		$MessageExpiringAccounts = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($ExpiringUsersTable.Count) -eq 1)
	{
		$MessageExpiringAccounts = "<i>There is $($ExpiringUsersTable.count) user that will expire in $ExpiringAccountDays days or less.</i></p>"
	}
	Else
	{
		$MessageExpiringAccounts = "<i>There are $($ExpiringUsersTable.count) users that will expire in $ExpiringAccountDays days or less.</i></p>"
	}
	
	
	If ($InactiveUsersTable -like "*Information: No Users*")
	{
		$MessageInactiveUsers = "<i>There are no users that is inactive.</i></p>"
		
	}
	ElseIf ($InactiveUsersTable -like "*Information: No Data*")
	{
		$MessageInactiveUsers = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($InactiveUsersTable.Count) -eq 1)
	{
		$MessageInactiveUsers = "<i>There is $($InactiveUsersTable.count) user that is inactive.</i></p>"
	}
	Else
	{
		$MessageInactiveUsers = "<i>There are $($InactiveUsersTable.count) users that are inactive.</i></p>"
	}
	
	
	If ($PasswordNeverExpiresTable -like "*Information: No Users*")
	{
		$MessagePasswordNeverExpires = "<i>There are no users with their password set to never expire.</i></p>"
		
	}
	ElseIf ($PasswordNeverExpiresTable -like "*Information: No Data*")
	{
		$MessagePasswordNeverExpires = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($PasswordNeverExpiresTable.Count) -eq 1)
	{
		$MessagePasswordNeverExpires = "<i>There is $($PasswordNeverExpiresTable.count) user with their password set to never expire.</i></p>"
	}
	Else
	{
		$MessagePasswordNeverExpires = "<i>There are $($PasswordNeverExpiresTable.count) users with their password set to never expire.</i></p>"
	}
	
	
	If ($NewCreatedUsersTable -like "*Information: No New Users*")
	{
		$MessageNewCreatedUsers = "<i>There are no newly created users in $UserCreatedDays Days or Less.</i></p>"
		
	}
	ElseIf ($NewCreatedUsersTable -like "*Information: No Data*")
	{
		$MessageNewCreatedUsers = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($NewCreatedUsersTable.Count) -eq 1)
	{
		$MessageNewCreatedUsers = "<i>There is 1 user that was created in $UserCreatedDays Days or Less.</i></p>"
	}
	Else
	{
		$MessageNewCreatedUsers = "<i>There are $($NewCreatedUsersTable.count) users that have been created in $UserCreatedDays Days or Less.</i></p>"
	}
	
	
	If ($UsersNoManagersTable -like "*Information: No New Users*")
	{
		$MessageUsersNoManager = "<i>There are no users without a manager set.</i></p>"
		
	}
	ElseIf ($UsersNoManagersTable -like "*Information: No Data*")
	{
		$MessageUsersNoManager = "<i>No Data; Report did not run.</i></p>"
	}
	ElseIf (($UsersNoManagersTable.Count) -eq 1)
	{
		$MessageUsersNoManager = "<i>There is 1 user does not have a manager set.</i></p>"
	}
	Else
	{
		$MessageUsersNoManager = "<i>There are $($UsersNoManagersTable.count) users that do not have managers set.</i></p>"
	}
	
	
	$mailmessage.Body = "<p>Good Morning,</p>

<p>I have finished running the daily automation reports, please review the results below. I have also attached an interactive HTML report to the e-mail for your convinence.</p>
<br>
<hr> </hr>
<p><h3>Report: Users with Passwords Expiring in $expireindays days or less</h3>
<br>
$MessageTable
<br>
<p>$daysalreadyq
<br>
$daysuntilq
<br>
$HTMLTable_PasswordExpiringNoEmail</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Users Expiring in $ExpiringAccountDays Days or less</h3>
<br>
$MessageExpiringAccounts
<br>
<p>$HTMLTable_ExpiringUsers</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Locked Out Users</h3>
<br>
$MessageLockedOut
<br>
<p>$HTMLTable_LockedOutUsers</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Inactive Users</h3>
<br>
$MessageInactiveUsers
<br>
<p>$HTMLTable_InactiveUsers</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Users With Password Set to Never Expire</h3>
<br>
$MessagePasswordNeverExpires
<br>
<p>$HTMLTable_PasswordNeverExpires</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Accounts Created in $UserCreatedDays Days or Less</h3>
<br>
$MessageNewCreatedUsers
<br>
<p>$HTMLTable_NewCreatedUsersTable</p>
<br>
<hr> </hr>
<br>
<p><h3>Report: Users Without a Manager Set</h3>
<br>
$MessageUsersNoManager
<br>
<p>$HTMLTable_NoManagerUsersTable</p>
<br>
<hr> </hr>
<br>
<p>Thanks and have a great $((get-date).dayofweek)!
<br>
<img src='https://www.thelazyadministrator.com/wp-content/uploads/2018/08/ewwewewe.png' alt='The Lazy Administrator'></p>
"
	(get-date -Format hh:mm:ss) + (": Attaching the HTML report to the email") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Attaching HTML Report"
	$mailmessage.Attachments.Add($att)
	
	(get-date -Format hh:mm:ss) + (": Sending the email") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Sending Email"
	Try
	{
		$smtpclient.Send($mailmessage)
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
	
	(get-date -Format hh:mm:ss) + (": Removing the attatchment") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	$att.Dispose()
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Summary email is not requested") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
}

(get-date -Format hh:mm:ss) + (": Checking to see if Teams summary is requested") | Out-File ($DirPath + "\" + "Log.txt") -Append

If ($SendSummaryTeamMessage -eq $True)
{
	(get-date -Format hh:mm:ss) + (": Teams summary message is requested") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	(get-date -Format hh:mm:ss) + (": Creating Teams Summary message") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Creating Teams Summary Message..."
	#Image on the left hand side, here I have a regular user picture
	$ItemImage = 'http://irisinstitute.ca/wp-content/uploads/sites/2/2017/11/Review.png'
	
	$ArrayTable = New-Object 'System.Collections.Generic.List[System.Object]'
	
	$Section = @{
		activityTitle = "Quick Glance"
		activityText  = "Report ran at $((Get-Date).ToString("hh:mm tt"))"
		activityImage = $ItemImage
		facts		  = @(
			@{
				name = "Passwords Expire Under $($expireindays) Days:"
				value = $Int_LessThan
			},
			@{
				name  = 'Passwords Expiring Today:'
				value = $Int_Today
			},
			@{
				name  = 'PW Expiring No Email Address:'
				value = $Int_NoEmail
			},
			@{
				name = "Users Expiring Under $($ExpiringAccountDays) Days:"
				value = $Int_ExpiringAccounts
			},
			@{
				name  = 'Locked Out Users:'
				value = $Int_LockedOut
			},
			@{
				name  = 'Inactive Users:'
				value = $Int_InactiveUsers
			},
			@{
				name  = 'Password Never Expire:'
				value = $Int_PWNeverExpires
			},
			@{
				name  = 'Newly Created Accounts:'
				value = $Int_NewAccounts
			},
			@{
				name  = 'Users No Managers:'
				value = $Int_NoManager
			}
		)
		
	}
	$ArrayTable.add($section)
	
	$body = ConvertTo-Json -Depth 8 @{
		title = "Daily Overview Report"
		text  = "I have finished running the daily automation reports, please review the results below. I have also sent a summary e-mail to $SummaryEmailAddress"
		sections = $ArrayTable
		
	}
	(get-date -Format hh:mm:ss) + (": Sending Summary Overview to Teams POST") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
	Write-Host "Sending Summary Overview to Teams POST" -ForegroundColor Green
	Try
	{
		Invoke-RestMethod -uri $SummaryTeamWebhookURL -Method Post -body $body -ContentType 'application/json' | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
	Catch
	{
		$_ | Out-File ($DirPath + "\" + "Log.txt") -Append
	}
	
}
Else
{
	(get-date -Format hh:mm:ss) + (": Teams Message not requested") | Out-File ($DirPath + "\" + "Log.txt") -Append
	
}
(get-date -Format hh:mm:ss) + (": FINISHED") | Out-File ($DirPath + "\" + "Log.txt") -Append
(get-date -Format hh:mm:ss) + (": $totalusers") | Out-File ($DirPath + "\" + "Log.txt") -Append


Write-Host "Done!"



