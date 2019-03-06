

$ExchangeServer = 'exchange.corp.com'
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri http://$ExchangeServer/PowerShell/ 
Import-PSSession $Session -DisableNameChecking -AllowClobber

$FirstName = @()
$MiddleInitial = @()
$LastName = @()
$UsernameSAMNoSpecial = @()
$FirstNameNoSpecial = @()
$LastNameNoSpecial = @()
$Global:PhoneExtension = @()
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
$global:TemplateOU = "DC=CORP,DC=COM"
$global:DepartmentsOU = 
$UserDepartment = @()

CLS
#------------------------------------------------Create username start-----------------------------------------#
#Gather users first name, required input and must not be empty or null
$FirstName = (Read-Host -Prompt 'Please input the users first name.')
$FirstName = (Get-Culture).TextInfo.ToTitleCase($FirstName)

#Ensure that first name is not empty
while ([string]::IsNullOrWhiteSpace($FirstName)) {$FirstName = read-host 'You left the first name empty, please enter a first name.'}

#Gather users middle initial, required input and must not be empty or null and must only be one character
$MiddleInitial = (Read-Host -Prompt 'Please input the users middle initial.')
$MiddleInitial = (Get-Culture).TextInfo.ToTitleCase($MiddleInitial)

#Ensure that middle initial isn't not more than 1 character or empty
while ([string]::IsNullOrWhiteSpace($MiddleInitial) -or ($MiddleInitial.Length -gt 1)) {$MiddleInitial = read-host 'You left the middle initial empty or input more than one character.'}

#Gather users last name, required input and must not be empty or null
$LastName = (Read-Host -Prompt 'Please input the users last name.')
$LastName = (Get-Culture).TextInfo.ToTitleCase($LastName)

#Ensure that last name is not empty
while ([string]::IsNullOrWhiteSpace($LastName)) {$LastName = read-host 'You left the last name empty, please enter a last name.'}

#Gathers user phone extension
CLS
$Global:PhoneExtension = (Read-Host -Prompt 'Please input the users 4 digit exension, leave blank if a user has not had a extension assigned')

function PhoneCheck {
if ($Global:PhoneExtension -match '[0-9][0-9][0-9][0-9]' -and ($Global:PhoneExtension.Length -eq 4))
    {
    }
elseif ([string]::IsNullOrWhiteSpace($Global:PhoneExtension))
{
$Global:PhoneExtension = $null
}
else
    {
$Global:PhoneExtension = $null 
$Global:PhoneExtension = Read-Host -Prompt 'Please input the users 4 digit exension, leave blank if a user has not had a extension assigned'
PhoneCheck
    }
}
PhoneCheck



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
$UserNameDisplay = ($FirstName + " " + $LastName)

#Create User Principle Name
$UserPrincipleName = $UniqueUserName + "@" + $DomainName

#Check username does not exist, if it does add numbers
$UniqueUserName = $UsernameSAM
while (Get-ADUser -Filter "SamAccountName -like '$UniqueUserName'"){$UniqueUserName = $UsernameSAM + ++$i}

$UniqueNumberAdd = $i


#Create User Principle Name
$UserPrincipleName = $UniqueUserName + "@" + $DomainName

$UniqueDisplayName = $UserNameDisplay
while (Get-ADUser -Filter "Name -eq '$UniqueDisplayName'"){$UniqueDisplayName = $UserNameDisplay + $UniqueNumberAdd}

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
CLS

function ManagerCheck {
$UserManagerCheck = Get-ADUser -Filter "SamAccountName -like '$UserManager'"
#$UserManagerInside = Get-ADUser -Filter "SamAccountName -like '$UserManager'"
if ($UserManagerCheck = [string]::IsNullOrWhiteSpace($UserManagerCheck))
    {
      cls
      $global:UserManager = (Read-Host -Prompt "Users manager not found please check the manager username")
      $UserManagerCheck = $null
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
#$UserDepartment = (Read-Host -Prompt "What is the users department, for example IT.")
CLS

Write-Host "The available departments are are `r`n"
Get-ADGroup -Filter * -SearchBase $global:DepartmentOU | Select -ExpandProperty Name | Sort-Object -Property Name
$global:Dept = (Read-Host -Prompt "`r`nWhat department is the user part of")


function DeptCheck {
$DeptCheck = Get-ADGroup -Filter "cn -like '$Dept'"
if ($DeptCheck = [string]::IsNullOrWhiteSpace($DeptCheck))
    {
      cls
      Write-Host "The available departments are are `r`n"
      Get-ADGroup -Filter * -SearchBase $global:DepartmentOU| Sort-Object -Property Name | Select -ExpandProperty Name 
      $global:Dept = (Read-Host -Prompt "`r`nDepartment not found, please try again using full group name")
      DeptCheck
    }
else
    { 
        CLS
    }
}

DeptCheck

$Dept = (Get-Culture).TextInfo.ToTitleCase($Dept)
$Dept = $Dept.Replace("Department","")
#while ([string]::IsNullOrWhiteSpace($UserDepartment)) {$UserDepartment = Read-Host 'You did not put the user in a department, please input the department the user is part of.'}
$UserCompany = (Read-Host -Prompt "What company does the user work for, if you do not enter data it will default to $DefaultCompany, please press enter if this is correct.")
CLS
$UserManager = (Read-Host -Prompt "Who is the users direct supervisor, please use the managers username and not full name.")

#Check attribuites have been populated
#while ([string]::IsNullOrWhiteSpace($UserTitle)) {$UserTitle = Read-Host 'You left the users title empty, please input a title for this user.'}

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
New-ADUser -Name $UniqueDisplayName -DisplayName $UniqueUserName -SamAccountName $UniqueUserName -GivenName $FirstName -Surname $LastName -Initials $MiddleInitial -OfficePhone $PhoneExtension -StreetAddress $UserStreetAddress -City $UserCity -State $UserState -Description $JobDescription -PostalCode $UserZipCode -Country "US" -UserPrincipalName $UserPrincipleName -Title $JobDescription -Department $Dept -Company $UserCompany -Manager $UserManager
Write-Host "Creating user and mailbox, please be patient"
#Wait 20 seconds to make sure user creation completes and propegates
Start-Sleep -Seconds 20
#Attach mailbox to new user
Enable-Mailbox -Identity $UserPrincipleName
#Create new email address based on companies defaults
Set-Mailbox $UserPrincipleName -EmailAddresses @{add=$EmailAddressExtra} -EmailAddressPolicyEnabled $False 
#Set email retention policies
Set-Mailbox $UserPrincipleName -PrimarySmtpAddress $EmailAddressExtra -RetentionPolicy $RetentionPolicy
#Disable Active Sync
Set-CasMailbox -Identity $UserPrincipleName  -ActiveSyncEnabled $false
#Copy permissions from user templates
get-ADuser -identity $UserTemplateCopyFrom -properties memberof | select-object memberof -expandproperty memberof | Add-AdGroupMember -Members $UniqueUserName
#Adds user to Horizon View Entitlement
Add-ADGroupMember -Identity $ViewEnt -Members $UsernameSAM
Set-Aduser -Identity $UniqueUserName -ChangePasswordAtLogon $false

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

#----------------------------------------------------Create Home Drive End--------------------------------------------------#

#----------------------------------------------------Create Report Start----------------------------------------------------------#
$UserInfoArray = New-Object PSObject
$UserInfoArray | Add-Member -type NoteProperty -Name 'Username' -Value $UserPrincipleName
$UserInfoArray | Add-Member -type NoteProperty -Name 'Display Name' -Value $UniqueDisplayName
$UserInfoArray | Add-Member -type NoteProperty -Name 'First Name' -Value $FirstName
$UserInfoArray | Add-Member -type NoteProperty -Name 'Middle' -Value $MiddleInitial
$UserInfoArray | Add-Member -type NoteProperty -Name 'Last Name' -Value $LastName
$UserInfoArray | Add-Member -type NoteProperty -Name 'Email Address Primary' -Value $EmailAddressExtra
$UserInfoArray | Add-Member -type NoteProperty -Name 'Phone Ext' -Value $PhoneExtension
$UserInfoArray | Add-Member -type NoteProperty -Name 'Manager' -Value $UserManager
$UserInfoArray | Add-Member -type NoteProperty -Name 'Department' -Value $Dept
$UserInfoArray | Add-Member -type NoteProperty -Name 'View Entitlement Group' -Value $ViewEnt
$UserInfoArray | Add-Member -type NoteProperty -Name 'Template Used' -Value $UserTemplateCopyFrom
$UserInfoArray | Add-Member -type NoteProperty -Name 'Home Drive Location' -Value "$FileServer\$UniqueUserName"
$UserInfoArray | Out-GridView
#----------------------------------------------------Create Report End----------------------------------------------------------#
