################################################################################################################################################################
# Script to set admin password to never expire
# Script accepts 2 parameters from the command line
#
# adminAccount - Optional - Administrator login ID for the tenant we are querying (Domain A)
# adminPassword - Optional - Administrator login password for the tenant we are querying (Domain A)
#
# Script returns C:\users.csv spreadsheet containng user accounts
#
# .\users.ps1 [-adminAccount user@xxx.com] [-adminPassword 'xxx'] 
#
# #DB 4.10.2016
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
# Function userList
# PURPOSE
#    To export users from Hosted Exchange
# INPUT
#    none
# RETURN
#    Comma Separated Value file C:\users.csv Headers: UserPrincipalName, DisplayName, FirstName, LastName, Department
###############################################################################
function userList
{   
    #Change this variable to desired path
	$outputFile= 'C:\users.csv'   
		
	#Following command retreives list of licensed users to named file
	Get-MsolUser | Where-Object { $_.isLicensed -eq "TRUE" } | Select-Object UserPrincipalName, DisplayName, FirstName, LastName, Department | Export-Csv -Path $outputFile
	
}

#Main
Function Main {

		
	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call function to export data to csv file
	userList
		
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
