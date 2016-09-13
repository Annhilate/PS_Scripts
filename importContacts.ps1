################################################################################################################################################################
# Script to import contacts into Hosted Exchange
# Script accepts 2 parameters from the command line
#
# adminAccount - Optional - Administrator login ID for the tenant we are querying (Domain A)
# adminPassword - Optional - Administrator login password for the tenant we are querying (Domain A)
#
# Imports formatted csv file to Domain; change the value of $inputFile to where file is located
#
# .\importContacts.ps1 [-adminAccount user@xxx.com] [-adminPassword 'xxx'] 
#
# #DB 4.10.2016
################################################################################################################################################################

#Input parameters
Param(
	[Parameter(Position=0, Mandatory=$False, ValueFromPipeline=$true)]
    [string] $adminAccount,
	[Parameter(Position=1, Mandatory=$False, ValueFromPipeline=$true)]
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
# Function userList
# PURPOSE
#    To import contacts info to Hosted Exchange
# INPUT
#    csv file containging formatted contact info 
# RETURN
#    none
###############################################################################
function userList
{   
	# this is where to change path of input file
	$inputFile = 'C:\contacts.csv'   
		
	#Following command Imports contats Display Name, email address, Firstname , Last Name. Second command imports department
	Import-Csv C-Path $inputFile | %{New-MailContact -Name $_.DisplayName -DisplayName $_.DisplayName -ExternalEmailAddress $_.UserPrincipalName -FirstName $_.FirstName -LastName $_.LastName}
	$Contacts = Import-CSV -Path  $inputFile
	$contacts | ForEach {Set-Contact $_.DisplayName -Department $_.Department}
}

#Main
Function Main {

		
	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call function to import contact info
	userList
		
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
