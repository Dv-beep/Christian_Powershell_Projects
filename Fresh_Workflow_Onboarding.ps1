param(
    [Parameter(Mandatory = $true, Position = 0)]
    [string]$First,

    [Parameter(Mandatory = $true, Position = 1)]
    [string]$Last,

    [Parameter(Mandatory = $true, Position = 2)]
    [string]$Title,

    [Parameter(Mandatory = $true, Position = 3)]
    [string]$Department,

    [Parameter(Mandatory = $true, Position = 4)]
    [string]$TicketID,

    [Parameter(Mandatory = $true, Position = 5)]
    [string]$Manager,

    [Parameter(Mandatory = $true, Position = 6)]
    [string]$EmployeeType
)

# Function to parse and verify OU based on Freshdesk number
function ParseOU {
    param (
        [string]$OU,
        [int]$FreshNumber
    )

    $OUnum = ($OU.Split()[-1] -replace '[()]')

    if ($OUnum -like "*-*") {
        $range = $OUnum.Split('-')
        try {
            $lowerEnd = [int]$range[0]
            $higherEnd = [int]$range[1]
            return ($FreshNumber -ge $lowerEnd -and $FreshNumber -le $higherEnd)
        } catch {
            return $false
        }
    } else {
        try {
            return [int]$OUnum -eq $FreshNumber
        } catch {
            return $false
        }
    }
}

# Parse Freshdesk department number and name
[int]$FreshNumber, $FreshDept = $Department.Split("-", 2)

# Define OU paths
$adminOU = "OU=Administration,OU=Lundquist-Users,DC=rei,DC=edu"
$campusOU = "OU=Campus,OU=Lundquist-Users,DC=rei,DC=edu"
$LUsers = "OU=Lundquist-Users,DC=rei,DC=edu"

# Determine the base OU based on Freshdesk department number
if ($FreshNumber -lt 1050 -and $FreshNumber -gt 999) {
    $baseOU = $adminOU
} else {
    $baseOU = $campusOU
}

# Function to find OU based on department name and number
function FindOU {
    param (
        [string]$baseOU,
        [string]$deptName
    )

    $OUlist = Get-ADOrganizationalUnit -Filter ("Name -like '*{0}*'" -f $deptName) -SearchBase $baseOU -SearchScope OneLevel

    if ($OUlist.Count -gt 0) {
        $OUBase = $OUlist[0].DistinguishedName
        $OU = Get-ADOrganizationalUnit -Filter 'Name -like "*Users*"' -SearchBase $OUBase -SearchScope OneLevel
        if ($OU) {
            return $OU.DistinguishedName
        }
    }

    $OUlist = Get-ADOrganizationalUnit -Filter 'Name -like "*"' -SearchBase $baseOU -SearchScope OneLevel
    foreach ($OU in $OUlist) {
        if (ParseOU $OU.Name $FreshNumber) {
            return $OU.DistinguishedName
        }
    }

    return $LUsers
}

# Find and return the OU
$OU = FindOU -baseOU $baseOU -deptName $FreshDept

# Define user credentials
$Username = "FreshOrchestrator"
$Password = "vW65!j5@2U"

# Convert plain text password to secure string
$SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force

# Create PSCredential object
$UserCredential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $Username, $SecurePassword

# Define the connection URI for Exchange
$ConnectionUri = "http://ex-101.rei.edu/PowerShell/"

# Create a new PowerShell session to Exchange
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $ConnectionUri -Authentication Kerberos -Credential $UserCredential

# Import the session
Import-PSSession $Session -DisableNameChecking -AllowClobber

# Construct User Principal Name (UPN)
$UPN = "$First.$Last@lundquist.org"
$aduser1 = "$First.$Last"
$DisplayName = "$Last, $First"

# Adjust UPN if it exceeds 20 characters
if ($aduser1.Length -gt 20) {
    $Initial = $First.Substring(0, 1)
    $aduser2 = "$Initial.$Last"
    $UPN = "$aduser2@lundquist.org"
} else {
    $aduser2 = $aduser1
}

# Define the email address
$email = $Manager
$username = $email -split '@' | Select-Object -First 1

# Create a new remote mailbox
New-RemoteMailbox -First $First -Last $Last -Name "$Last, $First" -UserPrincipalName $UPN -PrimarySmtpAddress $UPN -ResetPasswordOnNextLogon $false -Password $SecurePassword

# Clean up the session
Remove-PSSession $Session

# Wait to allow AD replication or caching to update
Start-Sleep -Seconds 10

# Update AD user attributes
Set-ADUser -Identity $aduser1 -Title $Title -Description $TicketID -Company 'The Lundquist Institute' -Department $Department -DisplayName $DisplayName -Manager $Username

# Move object to correct OU
$GUID = (Get-ADUser -Identity $aduser1).ObjectGUID
Move-ADObject -Identity $GUID -TargetPath $OU
