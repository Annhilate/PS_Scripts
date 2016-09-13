################################################################################################################################################################
# Script to set admin password to never expire
# Script accepts 2 parameters from the command line
#
# adminAccount - Optional - Administrator login ID for the tenant we are querying 
# adminPassword - Optional - Administrator login password for the tenant we are querying 
#
# Note:
# To connect to and manage Office 365 with the Azure AD Module for PowerShell, there are two prerequisites needed.
#
# Microsoft Online Services Sign-In Assistant for IT Professionals RTW
# Azure Active Directory Module for Windows PowerShell (64-bit version)
#
# To run the script
# Open an elevated Active Directory Module for Windows PowerShell console (not Windows PowerShell) 
# .\noexpire.ps1 [-adminAccount user@xxx.com] [-adminPassword 'xxx'] 
#
# #DB 2.29.2016
################################################################################################################################################################

#Input parameters
Param(
	[Parameter(Position=0, Mandatory=$false, ValueFromPipeline=$true)]
    [string] $adminAccount,
	[Parameter(Position=1, Mandatory=$false, ValueFromPipeline=$true)]
    [string] $adminPassword	
	)

###############################################################################
# Function connectExchange
# PURPOSE
#    Connects to Exchange Online Remote PowerShell using the tenant credentials
# INPUT
#    Tenant Admin username and password.
# RETURN
#    None.
###############################################################################
function connectExchange
{   
	Param( 
		[Parameter(	Mandatory=$false,Position=0)]
		[String]$aAccount,
		[Parameter(	Mandatory=$false,Position=1)]
		[String]$aPassword
		)  
		
	  
	#Did they provide creds?  If not, ask them for it. 
	if (([string]::IsNullOrEmpty($aAccount) -eq $false) -and ([string]::IsNullOrEmpty($aPassword) -eq $false)) 
	{ 
		#Encrypt password for transmission to Office365
		$SecurePassword = ConvertTo-SecureString -AsPlainText $aPassword -Force       
		  
		#Build credentials object
		$O365CREDS  = New-Object System.Management.Automation.PSCredential $aAccount, $SecurePassword  
	} 
	else 
	{ 
		#Build credentials object  
		$O365CREDS  = Get-Credential 
	} 	
		
	#Create remote Powershell session
	$SESSION = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://ps.outlook.com/powershell -Credential $O365CREDS -Authentication Basic -AllowRedirection    	

	#Import the session
    Import-PSSession $Session 
	
	# Download specific commands to be run locally
	Connect-MsolService -Credential $O365CREDS
}

###############################################################################
# Function setPass
# PURPOSE
#    To set password of user to never expire
# INPUT
#    none
# RETURN
#    None.
###############################################################################
function setPass
{   
	   
	$acct = Read-Host "Enter the email address of user " 	
		
	#Following command removes password expiration (for admin account)
	Set-msoluser –UserPrincipalName $acct -PasswordNeverExpires $True
	
}

#Main
Function Main {

	#Reminder to run script in AD Module
	""
	"Make sure you are using Active Directory Module for Windows PowerShell console!!!"
	""
	
	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call setPass to set account password to never expire
	setPass 
	
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
