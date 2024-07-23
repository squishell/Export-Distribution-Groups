function CheckAndInstallModule {
    param ([string]$moduleName)
    if (-not (Get-Module -ListAvailable -Name $moduleName)) {
        $installModule = Read-Host -Prompt "The $moduleName module is not installed. Do you want to install it now? (yes/no)"
        if ($installModule -eq 'yes') {
            Install-Module -Name $moduleName -Scope CurrentUser -Force
        } else {
            Write-Host "The script cannot proceed without the $moduleName module. Exiting..."
            exit
        }
    }
}

CheckAndInstallModule -moduleName "ExchangeOnlineManagement"
CheckAndInstallModule -moduleName "ImportExcel"
Import-Module ExchangeOnlineManagement

function Connect-Exchange {
    do {
        try {
            Connect-ExchangeOnline -UserPrincipalName $UserPrincipalName -ShowProgress $true -ErrorAction Stop
            Clear-Host
            break
        } catch {
            Write-Host 'Failed to connect. Please check the permissions or connection settings.'
            Disconnect-ExchangeOnline -Confirm:$false
            if ((Read-Host 'Retry with a different account? (yes/no)') -eq 'no') {
                exit
            }
        }
    } while ($true)
}

function Find-Group {
    param ([string]$groupName)
    # Initialize as an empty array to safely use +=
    $possibleGroups = @()

    # Concatenate results to the array, whether they are null, one or many
    $possibleGroups += @(Get-DistributionGroup -Identity $groupName -ErrorAction SilentlyContinue)
    $possibleGroups += @(Get-UnifiedGroup -Identity $groupName -ErrorAction SilentlyContinue)

    switch ($possibleGroups.Count) {
        0 {
            Write-Host "No groups found with the name $groupName. Please check the name and try again."
            return $null
        }
        1 {
            return $possibleGroups[0]
        }
        default {
            Write-Host "Multiple groups found. Please select the correct group:"
            for ($i = 0; $i -lt $possibleGroups.Count; $i++) {
                Write-Host "$($i+1). $($possibleGroups[$i].Name) - $($possibleGroups[$i].RecipientTypeDetails)"
            }
            $selectionOut = 0
            do {
                $selection = Read-Host 'Enter the number of the correct group'
                if ([int]::TryParse($selection, [ref]$selectionOut) -and $selectionOut -ge 1 -and $selectionOut -le $possibleGroups.Count) {
                    return $possibleGroups[$selectionOut - 1]
                }
                Write-Host "Invalid selection. Please enter a valid number."
            } while ($true)
        }
    }
}

function Get-GroupMembers {
   param (
     [string]$groupIdentity,
     [string]$recipientTypeDetails
   )

    #Logging for Error Check
    Write-Host "Group Identity: $groupIdentity" 
    Write-Host "Recipient Type Details: $recipientTypeDetails"

   # Retrieve members with more flexibility
   if ($recipientTypeDetails -eq "Office365Group" -or $recipientTypeDetails -eq "GroupMailbox") {
       $members = Get-UnifiedGroupLinks -Identity $groupIdentity -LinkType Members -ResultSize Unlimited
   } elseif ($recipientTypeDetails -eq "MailUniversalDistributionGroup") { 
       $members = Get-DistributionGroupMember -Identity $groupIdentity -ResultSize Unlimited
   } else {
       Write-Host "Unsupported group type: $($recipientTypeDetails)"
       return $null
   }
   return $members
 }

function Show-GroupMembers {
    param (
        [array]$members
    )
    Write-Host "`nMembers List:"
    foreach ($member in $members) {
        Write-Host "$($member.DisplayName)" 
        # Add to include email in Member List
        #- $($member.PrimarySmtpAddress)
    }
}

function Export-GroupMembers {
    param (
        [string]$groupName,
        [array]$members
    )
    $directoryPath = $PWD
    $fileName = "$groupName.csv"
    $memberDetails = @()
    $totalMembers = $members.Count
    
    Write-Host "Starting export of $totalMembers members..."

    for ($i = 0; $i -lt $totalMembers; $i++) {
        $member = $members[$i]
        if ([string]::IsNullOrWhiteSpace($member.PrimarySmtpAddress)) {
            Write-Host "Skipping user $($member.DisplayName) due to missing email address..."
            continue
        }
        try {
            $recipientDetails = Get-Recipient -Identity $member.PrimarySmtpAddress
            $memberDetails += [PSCustomObject]@{
                Name  = $recipientDetails.DisplayName
                Email = $recipientDetails.PrimarySmtpAddress
            }
            Write-Host "`rExporting user $(($i+1)) of $totalMembers..."
        } catch {
            Write-Host "Error processing $($member.DisplayName): $_"
        }
    }
    Write-Host "Exporting data to CSV..."
    $memberDetails | Export-Csv -Path "$directoryPath\$fileName" -NoTypeInformation
    Write-Host "Export completed successfully. File saved to $directoryPath\$fileName"
}

# Start of script execution
Connect-Exchange

do {
    $groupName = Read-Host -Prompt 'Please enter the name of the distribution group/list/M365 group'
    $group = Find-Group -groupName $groupName

    if ($group -eq $null) {
        Write-Host "Please try again."
        continue
    }

    Write-Host "`nFound Group:"
    Write-Host "Name: $($group.Name)"
    Write-Host "Email Address: $($group.PrimarySmtpAddress)`n"

    $confirmation = Read-Host -Prompt 'Is this the correct group? (yes/no)'

    if ($confirmation -eq 'yes') {
        $members = Get-GroupMembers -groupIdentity $group.Identity -recipientTypeDetails $group.RecipientTypeDetails
        if ($members) {
            Show-GroupMembers -members $members
            $export = Read-Host -Prompt 'Do you want to export the members to a CSV file? (yes/no)'
            if ($export -eq 'yes') {
                Export-GroupMembers -groupName $group.Name -members $members
            }
        } else {
            Write-Host "Failed to retrieve members or no members in group."
        }
    } else {
        Write-Host "Start over and select the correct group."
    }

    $repeat = Read-Host -Prompt 'Do you want to query another group? (yes/no)'
} while ($repeat -eq 'yes')


    
function Disconnect-Exchange {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [bool]
        $Confirm = $false
    )
    Disconnect-ExchangeOnline -Confirm:$Confirm
}

# Usage at the end of your script
Disconnect-Exchange -Confirm:$false
