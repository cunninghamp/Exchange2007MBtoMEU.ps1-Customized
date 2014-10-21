<#
.SYNOPSIS
Exchange2007MBtoMEU.ps1
#>

#requires -version 2

[CmdletBinding()]
param (
	[Parameter( Mandatory=$true)]
	[string]$DomainController
	)


#...................................
# Variables
#...................................

$now = Get-Date

$logfile = "Exchange2007MBtoMEU.log"


#...................................
# Logfile Strings
#...................................

$logstring0 = "====================================="
$logstring1 = " Exchange2007MBtoMEU.ps1"
$logstring2 = "You must supply a value for the -DomainController switch."
$logstring2 = "The cloud.csv file was not found in the current directory."


#...................................
# Initialization Strings
#...................................

$initstring0 = "Initializing..."
$initstring1 = "Loading the Exchange Server PowerShell snapin"
$initstring2 = "The Exchange Server PowerShell snapin did not load."
$initstring3 = "Setting scope to entire forest"


#...................................
# Functions
#...................................

#This function is used to write the log file if -Log is used
Function Write-Logfile()
{
	param( $logentry )
	$timestamp = Get-Date -DisplayHint Time
	"$timestamp $logentry" | Out-File $logfile -Append
}


function Main()
{
	#Script Logic flow
	#1. Pull User Info from cloud.csv file in the current directory
	#2. Lookup AD Info (DN, mail, proxyAddresses, and legacyExchangeDN) using the SMTP address from the CSV file
	#3. Save existing proxyAddresses
	#4. Add existing legacyExchangeDN's to proxyAddresses
	#5. Delete Mailbox
	#6. Mail-Enable the user using the cloud email address as the targetAddress
	#7. Disable RUS processing
	#8. Add proxyAddresses and mail attribute back to the object
	#9. Add msExchMailboxGUID from cloud.csv to the user object (for offboarding support)
	
	if($DomainController -eq [String]::Empty)
	{
		Write-Host $logstring2 -ForegroundColor Red
        Write-Logfile $logstring2
		Exit
	}
	
    if (Test-Path ".\cloud.csv")
    {	
        $CSVInfo = Import-Csv ".\cloud.csv"
    }
    else
    {
        Write-Host $logstring3 -ForegroundColor Red
        Write-Logfile $logstring3
		Exit
    }
	
    foreach($User in $CSVInfo)
	{
		$tmpstring = "Processing user $($User.OnPremiseEmailAddress)"
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring

        $tmpstring = "Calling LookupADInformationFromSMTPAddress"
		Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring

		$UserInfo = LookupADInformationFromSMTPAddress($User)
		
		#Check existing proxies for On-Premise and Cloud Legacy DN's as x500 proxies.  If not present add them.
		$CloudLegacyDNPresent = $false
		$LegacyDNPresent = $false
		foreach($Proxy in $UserInfo.ProxyAddresses)
		{
			if(("x500:$UserInfo.CloudLegacyDN") -ieq $Proxy)
			{
				$CloudLegacyDNPresent = $true
			}
			if(("x500:$UserInfo.LegacyDN") -ieq $Proxy)
			{
				$LegacyDNPresent = $true
			}
		}
		if(-not $CloudLegacyDNPresent)
		{
			$X500Proxy = "x500:" + $UserInfo.CloudLegacyDN
			
            $tmpstring = "Adding $X500Proxy to EmailAddresses"
            Write-Host $tmpstring -ForegroundColor Green
            Write-Logfile $tmpstring

			$UserInfo.ProxyAddresses += $X500Proxy
		}
		if(-not $LegacyDNPresent)
		{
			$X500Proxy = "x500:" + $UserInfo.LegacyDN

            $tmpstring = "Adding $X500Proxy to EmailAddresses"
			Write-Host $tmpstring -ForegroundColor Green
            Write-Logfile $tmpstring

			$UserInfo.ProxyAddresses += $X500Proxy
		}
		
		#Disable Mailbox
		$tmpstring = "Disabling Mailbox"
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		Disable-Mailbox -Identity $UserInfo.OnPremiseEmailAddress -DomainController $DomainController -Confirm:$false
		
		#Mail Enable
		$tmpstring = "Enabling Mailbox"
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		Enable-MailUser  -Identity $UserInfo.Identity -ExternalEmailAddress $UserInfo.CloudEmailAddress -DomainController $DomainController
		
		#Disable RUS
        $tmpstring = "Disabling RUS"		
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		Set-MailUser -Identity $UserInfo.Identity -EmailAddressPolicyEnabled $false -DomainController $DomainController
		
		#Add Proxies and Mail
		$tmpstring = "Adding EmailAddresses and WindowsEmailAddress"
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		Set-MailUser -Identity $UserInfo.Identity -EmailAddresses $UserInfo.ProxyAddresses -WindowsEmailAddress $UserInfo.Mail -DomainController $DomainController
		
		#Set Mailbox GUID.  Need to do this via S.DS as Set-MailUser doesn't expose this property.
		$ADPath = "LDAP://" + $DomainController + "/" + $UserInfo.DistinguishedName
		$ADUser = New-Object -TypeName System.DirectoryServices.DirectoryEntry -ArgumentList $ADPath
		$MailboxGUID = New-Object -TypeName System.Guid -ArgumentList $UserInfo.MailboxGUID
		[Void]$ADUser.psbase.invokeset('msExchMailboxGUID',$MailboxGUID.ToByteArray())
		
        $tmpstring = "Setting Mailbox GUID $($UserInfo.MailboxGUID)"
        Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		
        $ADUser.psbase.CommitChanges()
		
        $tmpstring = "Migration Complete for $($UserInfo.OnPremiseEmailAddress)"
		Write-Host $tmpstring -ForegroundColor Green
        Write-Logfile $tmpstring
		Write-Host ""
		Write-Host ""
	}
}

function LookupADInformationFromSMTPAddress($CSV)
{
	$Mailbox = Get-Mailbox $CSV.OnPremiseEmailAddress -ErrorAction SilentlyContinue
	
	if($Mailbox -eq $null)
	{
		Write-Host "Get-Mailbox failed for" $CSV.OnPremiseEmailAddress -ForegroundColor Red
		continue
	}
	
	$UserInfo = New-Object System.Object
	
	$UserInfo | Add-Member -Type NoteProperty -Name OnPremiseEmailAddress -Value $CSV.OnPremiseEmailAddress
	$UserInfo | Add-Member -Type NoteProperty -Name CloudEmailAddress -Value $CSV.CloudEmailAddress
	$UserInfo | Add-Member -Type NoteProperty -Name CloudLegacyDN -Value $CSV.LegacyExchangeDN
	$UserInfo | Add-Member -Type NoteProperty -Name LegacyDN -Value $Mailbox.LegacyExchangeDN
	$ProxyAddresses = @()
	foreach($Address in $Mailbox.EmailAddresses)
	{
		$ProxyAddresses += $Address
	}
	$UserInfo | Add-Member -Type NoteProperty -Name ProxyAddresses -Value $ProxyAddresses
	$UserInfo | Add-Member -Type NoteProperty -Name Mail -Value $Mailbox.WindowsEmailAddress
	$UserInfo | Add-Member -Type NoteProperty -Name MailboxGUID -Value $CSV.MailboxGUID
	$UserInfo | Add-Member -Type NoteProperty -Name Identity -Value $Mailbox.Identity
	$UserInfo | Add-Member -Type NoteProperty -Name DistinguishedName -Value (Get-User $Mailbox.Identity).DistinguishedName
	
	$UserInfo
}


#...................................
# Initialize
#...................................

#Log file is overwritten each time the script is run to avoid
#very large log files from growing over time

$timestamp = Get-Date -DisplayHint Time
"$timestamp $logstring0" | Out-File $logfile
Write-Logfile $logstring1
Write-Logfile "  $now"
Write-Logfile $logstring0


Write-Host $initstring0
Write-Logfile $initstring0

#Add Exchange 2007 snapin if not already loaded in the PowerShell session
if (!(Get-PSSnapin | where {$_.Name -eq "Microsoft.Exchange.Management.PowerShell.Admin"}))
{
	Write-Verbose $initstring1
	Write-Logfile $initstring1
	try
	{
		Add-PSSnapin Microsoft.Exchange.Management.PowerShell.Admin -ErrorAction STOP
	}
	catch
	{
		#Snapin was not loaded
		Write-Verbose $initstring2
		Write-Logfile $initstring2
		Write-Warning $_.Exception.Message
		EXIT
	}
}


#...................................
# Script
#...................................


Main


#...................................
# End
#...................................


$timestamp = Get-Date -DisplayHint Time
"$timestamp $logstring0" | Out-File $logfile -Append
Write-Logfile $logstring1
Write-Logfile "  $now"
Write-Logfile $logstring0
