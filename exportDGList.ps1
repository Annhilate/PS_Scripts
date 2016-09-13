################################################################################################################################################################
# Script to Download Distribution Groups and members
# Script accepts 2 parameters from the command line
#
# adminAccount - Optional - Administrator login ID for the tenant we are querying (Domain A)
# adminPassword - Optional - Administrator login password for the tenant we are querying (Domain A)
#
# Imports formatted csv file to Domain
#
# .\exportDGList.ps1 [-adminAccount user@xxx.com] [-adminPassword 'xxx'] 
#
# #DB 4.12.2016
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
#    Tenant Admin username and password, if none provided method will ask for credentials
# RETURN
#    None
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
# Function Distrbution Group Export
# PURPOSE
#    To export Distribution Group Lists and Members
# INPUT
#    none
# RETURN
#    csv file with Dist. Group info
###############################################################################
function distGroup
{   
$OutputFile = 'C:\DistributionGroupMembers.csv'   #The CSV Output file that is created, change for your purposes  

#Prepare Output file with headers  
Out-File -FilePath $OutputFile -InputObject 'Distribution Group DisplayName,Distribution Group Email,Member DisplayName, Member Email, Member Type' -Encoding ASCII 
  
#Get all Distribution Groups from Office 365  
$objDistributionGroups = Get-DistributionGroup -ResultSize Unlimited  
  
#Iterate through all groups, one at a time      
Foreach ($objDistributionGroup in $objDistributionGroups)  
{      
     
    write-host 'Processing $($objDistributionGroup.DisplayName)...'  
  
    #Get members of this group  
    $objDGMembers = Get-DistributionGroupMember -Identity $($objDistributionGroup.PrimarySmtpAddress)  
      
    write-host 'Found $($objDGMembers.Count) members...'  
      
    #Iterate through each member  
    Foreach ($objMember in $objDGMembers)  
    {  
        
		#Out-File -FilePath $OutputFile -InputObject $($objDistributionGroup.DisplayName), $($objDistributionGroup.PrimarySMTPAddress), $($objMember.DisplayName), $($objMember.PrimarySMTPAddress), $($objMember.RecipientType) -Encoding ASCII -append
		#$newObj=New-Object PSObject 
		$newLine = "{0},{1},{2},{3},{4}" -f $($objDistributionGroup.DisplayName), $($objDistributionGroup.PrimarySMTPAddress), $($objMember.DisplayName), $($objMember.PrimarySMTPAddress), $($objMember.RecipientType)
		$newLine | Add-Content -Path $OutputFile
		write-host $($objDistributionGroup.DisplayName),$($objDistributionGroup.PrimarySMTPAddress),$($objMember.DisplayName),$($objMember.PrimarySMTPAddress),$($objMember.RecipientType)
		
	}

	 
}  
}

#Main
Function Main {

		
	#Remove all existing Powershell sessions
	Get-PSSession | Remove-PSSession
	
	#Call ConnectTo-ExchangeOnline function with correct credentials
	connectExchange -aAccount $adminAccount -aPassword $adminPassword
	
	#Call function to import contact info
	distGroup
		
	#Clean up session
	Get-PSSession | Remove-PSSession
}

# Start script
. Main
