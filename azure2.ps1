################################################################################################################################################################
# Azure Rights Management Activation
# Script accepts 3 parameters from the command line
#
# adminAccount - Mandatory - Administrator login ID for the tenant we are querying
# adminPassword - Mandatory - Administrator login password for the tenant we are querying
# mailbox - Mandatory - User account that has a mailbox
##
# To run the script
#
# .\azure.ps1 -adminAccount etccland@xxx.com -adminPassword 'xxx' -mailbox user@xxx.com
#
# #DB 6.10.2015
################################################################################################################################################################

#Input parameters
Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $adminAccount,
	[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $adminPassword,	
	[Parameter(Position=2, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $mailbox
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
		[Parameter(	Mandatory=$true,Position=0)]
		[String]$aAccount,
		[Parameter(	Mandatory=$true,Position=1)]
		[String]$aPassword

    )  
		
	#Encrypt password for transmission to Office365
	$SecurePassword = ConvertTo-SecureString -AsPlainText $aPassword -Force    
	
	#Build credentials object
	$Office365Credentials  = New-Object System.Management.Automation.PSCredential $aAccount, $SecurePassword
	
	#Create remote Powershell session
	$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Office365Credentials -Authentication Basic -AllowRedirection    	

	#Import the session
    Import-PSSession $Session 
}

###############################################################################
# Function configIRM
# PURPOSE
#    Sets and Tests correct IRM Configuration
# INPUT
#    mail account
# RETURN
#    None.
###############################################################################
function configIRM
{   
	Param( 
		[Parameter(	Mandatory=$true,Position=0)]
		[String]$mailUser
    )  
		
	#Set IRM Configuration for North America
	Set-IRMConfiguration -RMSOnlineKeySharingLocation "https://sp-rms.na.aadrm.com/TenantManagement/ServicePartner.svc"   
	
	#Import Trusted Domain
	Import-RMSTrustedPublishingDomain -RMSOnline -name "RMS Online"
	
	#Test with mailbox account
	Test-IRMConfiguration -Sender $mailUser
	
	#Disable Client Access Server
    Set-IRMConfiguration -ClientAccessServerEnabled $false
	
	#Enable Internal Licensing
	Set-IRMConfiguration -InternalLicensingEnabled $true
	
	#Test with mailbox account
	Test-IRMConfiguration -Sender $mailUser

}

#Main
Function Main {

	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call configIRM to setup and test IRM Configuration
	configIRM -mailUser $mailbox
	
	#Call configIRM to setup and test IRM Configuration Second Time
	configIRM -mailUser $mailbox
	
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
