################################################################################################################################################################
# Script to allow Domain A to ‘federate’ and share calendars with Domain B
# Script accepts 4 parameters from the command line
#
# adminAccount - Mandatory - Administrator login ID for the tenant we are querying (Domain A)
# adminPassword - Mandatory - Administrator login password for the tenant we are querying (Domain A)
# domB - Mandatory - Domain that we are establishing trust with (Domain B)
# domBName - Mandatory - Name of company (Domain B)
##
# To run the script
#
# .\federate.ps1 -adminAccount user@xxx.com -adminPassword 'xxx' -domB 'xxx.com' -domBName 'XXX' 
#
# #DB 2.29.2016
################################################################################################################################################################

#Input parameters
Param(
	[Parameter(Position=0, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $adminAccount,
	[Parameter(Position=1, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $adminPassword,	
	[Parameter(Position=2, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $domB,
	[Parameter(Position=3, Mandatory=$true, ValueFromPipeline=$true)]
    [string] $domBName
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
# Function enableShare
# PURPOSE
#    To enable sharing of Calendars between tenants (this is one way only)
# INPUT
#    Name of remote Domain, Company Name
# RETURN
#    None.
###############################################################################
function enableShare
{   
	Param( 
		[Parameter(	Mandatory=$true,Position=0)]
		[String]$dom,
		[Parameter(	Mandatory=$true,Position=1)]
		[String]$name

    )   
		
	#Allow customizations to occur on tenant
	Enable-OrganizationCustomization  
	
	#The following command will establish the relationship from your tenant (Domain A) to the remote tenant (Domain B), but is only one way
	Get-FederationInformation -DomainName $dom | New-OrganizationRelationship -Name $name -FreeBusyAccessEnabled $true -FreeBusyAccessLevel LimitedDetails
	
}

#Main
Function Main {

	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call enableShare to setup federation
	enableShare -dom $domB -name $domBName
	
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
