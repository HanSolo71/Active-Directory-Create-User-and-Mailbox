

$ExchangeServer = 'exchange.corp.com'
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ 
Import-PSSession $Session -DisableNameChecking -AllowClobber

$FirstName = @()
$MiddleInitial = @()
$LastName = @()
$UsernameSAMNoSpecial = @()
$FirstNameNoSpecial = @()
$LastNameNoSpecial = @()
$PhoneExtension = @()
$Username = @()
$UniqueUserName = @()
$UserStreetAddress = @()
$UserCity = @()
$UserState = @()
$JobDescription = @()
$UserNameDisplay = @()
$UniqueDisplayName = @()
$UserPrincipleName = @()
$UserManagerCheck = @()
$UniqueNumberAdd = @()
$EmailAddress = @()
$EmailAlias = @()
$global:UserTemplateCopyFrom = @()
$global:UserTemplateCheck = @()
$global:UserTemplateCopyFrom = @()
$global:ViewEntCheck
$ViewEnt = @()
$Global:EntitlementsOU = "DC=CORP,DC=COM"
$RetentionPolicy = "Default - Delete all items except Notes over 3 years old"
$TempPassword = "Pa$$word1"
$PrimaryEmailDomain = "@example2.com"
$DomainName = "example.com"
$DefaultAddress = "SR 405 Kennedy Space Center"
$DefaultState = "FL"
$DefaultZip = "32899"
$DefaultCountry = "US"
$DefaultCity = "Cape Canaveral"
$DefaultCompany = "NASA"
$FileServer = "\\FileServer\H_Drives$"
$i = $Null
$global:UserManager = @()
$global:TemplateOU = "DC=CORP,DC=COM"

CLS
#------------------------------------------------Create username start-----------------------------------------#
#Gather users first name, required input and must not be empty or null
$FirstName = (Read-Host -Prompt 'Please input the users first name.')

#Ensure that first name is not empty
while ([string]::IsNullOrWhiteSpace($FirstName)) {$FirstName = read-host 'You left the first name empty, please enter a first name.'}

#Gather users middle initial, required input and must not be empty or null and must only be one character
$MiddleInitial = (Read-Host -Prompt 'Please input the users middle initial.')
(Get-Culture).TextInfo.ToTitleCase($MiddleInitial)

#Ensure that middle initial isn't not more than 1 character or empty
while ([string]::IsNullOrWhiteSpace($MiddleInitial) -or ($MiddleInitial.Length -gt 1)) {$MiddleInitial = read-host 'You left the middle initial empty or input more than one character.'}

#Gather users last name, required input and must not be empty or null
$LastName = (Read-Host -Prompt 'Please input the users last name.')

#Ensure that last name is not empty
while ([string]::IsNullOrWhiteSpace($LastName)) {$LastName = read-host 'You left the last name empty, please enter a last name.'}

#Gathers user phone extension, required input, mustn ot be empty or null, and must only user numbers
$PhoneExtension = (Read-Host -Prompt 'Please input the users 4 digit exension, numbers only')

#Ensure that phone extension is only 4 numbers
while ($PhoneExtension -notmatch '[0-9][0-9][0-9][0-9]') {$PhoneExtension = Read-Host -Prompt 'Please only use numbers in the phone extensione and ensure it is 4 characters.'}

#Ensure that phone extension is only 4 charcters long
while ($PhoneExtension.Length -ne 4) {$PhoneExtension = Read-Host -Prompt 'Please only use the 4 digit extension'}

#Ensure that phone extension is not empty
while ([string]::IsNullOrWhiteSpace($PhoneExtension)) {$PhoneExtension = read-host 'You left the phone extension empty, please input a 4 digit extension'}

#Set users description of their job, for example "Call Center Representative"
$JobDescription = (Read-Host -Prompt 'Please input a title for the users position, for example "Call Center Representative"')

#Ensure job description is not empty
while ([string]::IsNullOrWhiteSpace($JobDescription)) {$JobDescription = read-host 'You left the job description empty, please input the users job title.'}

#Create user name
$FirstNameNoSpecial = $Firstname -replace '[^\p{L}\p{Nd}]'
$LastNameNoSpecial = $LastName -replace '[^\p{L}\p{Nd}]'

if ($LastNameNoSpecial.Length -ge 6)
{ 
$UsernameSAM = $FirstNameNoSpecial.Substring(0,1) + $MiddleInitial + $LastNameNoSpecial.Substring(0,6)
}
else
{
$UsernameSAM = $FirstNameNoSpecial.Substring(0,1) + $MiddleInitial + $LastNameNoSpecial
}
$UsernameSAM = $UsernameSAM.ToLower()


#Create Display Username
(Get-Culture).TextInfo.ToTitleCase($FirstName + " " + $LastName)

#Create User Principle Name
$UserPrincipleName = $UniqueUserName + "@" + $DomainName

#Check username does not exist, if it does add numbers
CLS
$UniqueUserName = $UsernameSAM
while (Get-ADUser -Filter "SamAccountName -like '$UniqueUserName'"){$UniqueUserName = $UsernameSAM + ++$i}


$UniqueNumberAdd = $i

#Create User Principle Name
$UserPrincipleName = $UniqueUserName + "@" + $DomainName

$UserNameDisplay = $FirstName + " " + $LastName
$UniqueDisplayName = $UserNameDisplay
while (Get-ADUser -Filter 'Name -eq "$UniqueDisplayName"'){$UniqueDisplayName = $UserNameDisplay + $UniqueNumberAdd}
Write-Host "The new users username is $UniqueDisplayName"
Write-Host "The new users logon name is $UniqueUsername"

#--------------------------------------------Create Username End------------------------------------------------#



#--------------------------------------------Create user address start------------------------------------------#

#Get users Street Address, if the input is left empty then it will automatically default to 618 Kenmoor Ave SE
$UserStreetAddress = (Read-Host -Prompt "Please input the users street address, will default to $DefaultAddress, please press enter if this is correct")

#Get users city
$UserCity = (Read-Host -Prompt "Please input the users city, will default to $DefaultCity, please press enter if this is correct")

#Get users state
$UserState = (Read-Host -Prompt "Please input the users state initials only, if nothing is input it will default to' $DefaultState, please press enter if this is correct")

#Get user zip code
$UserZipCode = (Read-Host -Prompt "Please input the users ZIP code in 5 digit format, if left blank will default to $DefaultZip, please press enter if this is correct")


#Get users country
$UserCountry = (Read-Host -Prompt "Please enter two digit country code, if nothing is input this will default to $DefaultCountry, please press enter if this is correct")

#Ensure that user street address is not empty if it is, uses default address 
while ([string]::IsNullOrWhiteSpace($UserStreetAddress)) {$UserStreetAddress = $DefaultAddress}
#Ensure that user city is not empty, if it is uses default city
while ([string]::IsNullOrWhiteSpace($UserCity)) {$UserCity = $DefaultCity}
#Ensure that users state is not empty, if it is uses default state
while ([string]::IsNullOrWhiteSpace($UserState)) {$UserState = $DefaultState}
#Ensure that only two digit code for state is used
while ($UserState.Length -ne 2) {$UserState = Read-Host -Prompt 'Please only use the abbreviation for the State'}
#Ensure that zip code is not empty, if not uses default value
while ([string]::IsNullOrWhiteSpace($UserZipCode)) {$UserZipCode = $DefaultZip}
#Ensures that only 5 digit zip code is used
while ($UserZipCode.Length -ne 5) {$UserZipCode = Read-Host -Prompt 'Please only use the 5 digit ZIP code'}
#Ensure zip code only has numbers in it
while ($UserZipCode -notmatch '[0-9][0-9][0-9][0-9][0-9]') {$UserZipCode = Read-Host -Prompt 'Please only use numbers in the zip code'}
#Ensure country code is not empty, if it is use default country
while ([string]::IsNullOrWhiteSpace($UserCountry)) {$UserCountry = $DefaultCountry}
#Ensure that users country code is only 2 digits
while ($UserCountry.Length -ne 2) {$UserCountry = Read-Host -Prompt 'Please only use 2 digit country codes'}

#-----------------------------------------------------Create user address end-----------------------------------------#


#-----------------------------------------------------Create user organization attributes start-----------------------#
#Function checks for manager existence in active directory
#Function checks for manager existence in active directory


function ManagerCheck {
$UserManagerCheck = Get-ADUser -Filter "SamAccountName -like '$UserManager'"
#$UserManagerInside = Get-ADUser -Filter "SamAccountName -like '$UserManager'"
if ($UserManagerCheck = [string]::IsNullOrWhiteSpace($UserManagerCheck))
    {
      cls
      $global:UserManager = (Read-Host -Prompt "Users manager not found please check the manager username")
      ManagerCheck 
    }
else
    { 
        {continue}
        CLS
    }
}


#Gather organziational data
#$UserTitle = (Read-Host -Prompt "What is the users job title, for example Network Administrator.")
$UserDepartment = (Read-Host -Prompt "What is the users department, for example IT.")
$UserCompany = (Read-Host -Prompt "What company does the user work for, if you do not enter data it will default to $DefaultCompany, please press enter if this is correct.")
$UserManager = (Read-Host -Prompt "Who is the users direct supervisor, please use the managers username and not full name.")

#Check attribuites have been populated
#while ([string]::IsNullOrWhiteSpace($UserTitle)) {$UserTitle = Read-Host 'You left the users title empty, please input a title for this user.'}
while ([string]::IsNullOrWhiteSpace($UserDepartment)) {$UserDepartment = Read-Host 'You did not put the user in a department, please input the department the user is part of.'}
#Default company name if no input
while ([string]::IsNullOrWhiteSpace($UserCompany)) {$UserCompany = $DefaultCompany }
while ([string]::IsNullOrWhiteSpace($UserManager)) {$UserManager = Read-Host 'You left their manager empty, please input a manager username'}
#Run manager check function
ManagerCheck



#----------------------------------------------------Create user organization attributes end--------------------------#

#----------------------------------------------------Create user email start------------------------------------------------#
#Creates primary email address
$EmailAddress = $FirstName + $LastName.Substring(0,1) + $UniqueNumberAdd
#Create secondary email address
$EmailAddressExtra = $EmailAddress + $PrimaryEmailDomain



#----------------------------------------------------Create user email end--------------------------------------------------#



#----------------------------------------------------Copy permissions from template start-----------------------------------#
Write-Host "The available template users are `r`n"
Get-ADUser -Filter * -SearchBase $TemplateOU | Select -ExpandProperty SAMAccountName | Sort-Object -Property SAMAccountName
$global:UserTemplateCopyFrom = (Read-Host -Prompt "`r`nWhat template would you like to copy from, only accounts in the User Template OU will be accepted ")



function TemplateUserCheck {
$UserTemplateCheck = Get-ADUser -SearchBase $global:TemplateOU -Filter "SamAccountName -like '$UserTemplateCopyFrom'"
if ($UserTemplateCheck = [string]::IsNullOrWhiteSpace($UserTemplateCheck))
    {
      cls
      Write-Host "The available template users are $TemplateOU`r`n"
      Get-ADUser -Filter * -SearchBase $TemplateOU | Select -ExpandProperty SAMAccountName | Sort-Object -Property SAMAccountName
      $global:UserTemplateCopyFrom = (Read-Host -Prompt "User template not found in 'User Template OU'")
      $UserTemplateCheck = $null
      TemplateUserCheck  
    }
else
    {
      {continue}
      CLS
    }
}


TemplateUserCheck


#----------------------------------------------------Copy permissions from template end-------------------------------------#

#----------------------------------------------------Start Horizon View Entitlement-------------------------------------------#
CLS
Write-Host "The available entitlement groups are `r`n"
Get-ADGroup -Filter * -SearchBase $EntitlementsOU | Select -ExpandProperty Name | Sort-Object -Property Name
$ViewEnt = (Read-Host -Prompt "`r`nWhat Horizon View Entitlement group should the user be made part of ")

function AddViewEnt {
$ViewEntCheck = Get-ADGroup -Filter "cn -like '$ViewEnt'"
if ($ViewEntCheck = [string]::IsNullOrWhiteSpace($ViewEntCheck))
    {
      cls
      Write-Host "The available entitlement groups are `r`n"
      Get-ADGroup -Filter * -SearchBase $Global:EntitlementsOU | Sort-Object -Property Name | Select -ExpandProperty Name 
      $global:ViewEnt = (Read-Host -Prompt "`r`nHorizon View Entitlement group not found, please try again using full group name")
      AddViewEnt
    }
else
    { 
        CLS
    }
}

AddViewEnt
#-------------------------------------------End Horizon View Entitlement------------------------------------------------#

#----------------------------------------------------Create User Start------------------------------------------------------#

#Create user
New-ADUser -Name $UniqueDisplayName -DisplayName $UniqueDisplayName -SamAccountName $UniqueUserName -GivenName $FirstName -Surname $LastName -Initials $MiddleInitial -OfficePhone $PhoneExtension -StreetAddress $UserStreetAddress -City $UserCity -State $UserState -Description $JobDescription -PostalCode $UserZipCode -Country "US" -UserPrincipalName $UserPrincipleName -Title $JobDescription -Department $UserDepartment -Company $UserCompany -Manager $UserManager
Write-Host "Creating user and mailbox, this can take up to 40 seconds, please be patient"
Set-Aduser -Identity $UniqueUserName -ChangePasswordAtLogon $false
#Wait 20 seconds to make sure user creation completes and propegates
Start-Sleep -Seconds 20
#Attach mailbox to new user
Enable-Mailbox -Identity $UniqueDisplayName
#Create new email address based on companies defaults
Set-Mailbox $UniqueDisplayName -EmailAddresses @{add=$EmailAddressExtra} -EmailAddressPolicyEnabled $False 
#Set email retention policies
Set-Mailbox $UniqueDisplayName -PrimarySmtpAddress $EmailAddressExtra -RetentionPolicy $RetentionPolicy
#Disable Active Sync
Set-CasMailbox -Identity $UniqueDisplayName  -ActiveSyncEnabled $false
#Copy permissions from user templates
get-ADuser -identity $UserTemplateCopyFrom -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $UniqueUserName
#Adds user to Horizon View Entitlement
Add-ADGroupMember -Identity $ViewEnt -Members $UsernameSAM

#----------------------------------------------------Create User End--------------------------------------------------------#



#----------------------------------------------------Create Home Drive Start------------------------------------------------#

#Creating home directory and set permissions
$UniqueUserNameLower = $UniqueUserName.ToLower()
new-item -path "$FileServer\$UniqueUserNameLower" -ItemType Directory
$acl = get-acl "$FileServer\$UniqueUserName"
$FileSystemRights = [System.Security.AccessControl.FileSystemRights]"FullControl"
$AccessControlType = [System.Security.AccessControl.AccessControlType]::Allow
$InheritanceFlags = [System.Security.AccessControl.InheritanceFlags]"ContainerInherit, ObjectInherit"
$PropagationFlags = [System.Security.AccessControl.PropagationFlags]"InheritOnly"
$AccessRule = New-Object System.Security.AccessControl.FileSystemAccessRule ("$DomainName\$UniqueUserName", $FileSystemRights, $InheritanceFlags, $PropagationFlags, $AccessControlType)
$acl.AddAccessRule($AccessRule)
Set-Acl -Path "$FileServer\$UniqueUserName" -AclObject $acl -ea Stop

#----------------------------------------------------Create Home Drive End--------------------------------------------------#----#
