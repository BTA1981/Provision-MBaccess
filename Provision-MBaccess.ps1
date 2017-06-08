#Requires -Version 3.0
<#
.SYNOPSIS
    Read members of a certain group en add them to a shared mailbox.
    Script also needs be te able to update and remove users from a mailbox.

    Script is not fully tested.
    
.DESCRIPTION
    
    Prerequisites: Modules ActiveDirectory, GroupPolicy
.NOTES
  Version:        1.0
  Author:         Bart Tacken - Client ICT Groep
  Creation Date:  21-02-2017
  Purpose/Change: Initial script development
.EXAMPLE
    
#>

#---------------------------------------------------------[Initialisations]--------------------------------------------------------
#Set Error Action to Silently Continue
$ErrorActionPreference = 'SilentlyContinue'
[string]$DateStr = (Get-Date).ToString("s").Replace(":","-") # +"_" # Easy sortable date string
Start-Transcript ('c:\windows\temp\' + $DateStr  + '_Provision-MBaccess.log') # Start logging 
#----------------------------------------------------------[Declarations]----------------------------------------------------------
# Variables
If (!Get-Module activedirectory) { import-module activedirectory}
#$CSVpath = "C:\temp\iriszorg\GroupsAndMailboxesDemo.csv"
#$CSV = Import-CSV $CSVpath
$AccessRights = 'FullAccess'
$CurrentMemberArray = @() 
$CurrentDistributionGroupMembersArray = @()
$ComputerName = 'ps.outlook.com' 
 #----------------------------------------------------------[Functions]------------------------------------------------------------
 Function Connect-EXOnline {
    $URL = "https://ps.outlook.com/powershell"  
    $Credentials = Get-Credential -Message "Enter your Office 365 admin credentials"
    $EXOSession = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri $URL -Credential $Credentials -Authentication Basic -AllowRedirection -Name "Exchange Online"
        Import-PSSession $EXOSession
}
#-----------------------------------------------------------[Execution]------------------------------------------------------------
# Make connection with Exchange Online
$Session = Get-PSSession | where { $_.ComputerName -eq $ComputerName -and $_.State -eq 'Opened' }
if (!$Session) {
    try { Connect-EXOnline -EA 1 } catch { throw }
}

# First check if every shared mailbox contains a "PB_" security group
# Log all mailboxes that don't have one.
$AllSharedMailboxes = Get-mailbox -RecipientTypeDetails sharedmailbox -Resultsize unlimited | Select-Object -ExpandProperty name # Get current list of shared mailboxes
$AllSharedMailboxesWithPB_Group = @()

$ErrorActionPreference = 'stop'

ForEach ($MBname in $AllSharedMailboxes) {
    Try {
        Get-ADgroup ("PB_" + $MBname) -ErrorAction Stop
        $AllSharedMailboxesWithPB_Group += $MBName 
    }
    Catch {
        Write-Output "Security Group [$MBName] does not exist" | out-file "C:\temp\MailboxDoesNotExist.log" -Append
        Continue
    } # End Catch
} # End ForEAch

$ErrorActionPreference = 'silentlycontinue'  


# Go through all shared mailboxes and for each mailbox:
    # Extract all members of the corresponding security group
    # Add permissions for all these members to the mailbox
ForEach ($Mailbox in $AllSharedMailboxesWithPB_Group) {
    
    #If ($Mailbox -like "AchterhoekFACTBegeleiding") { # for testing purposes
        
        
    $MailboxString = (Get-Mailbox $Mailbox) | Select-Object -ExpandProperty name    
    #$MailboxString = $MailboxObject | Select-Object -ExpandProperty name
          
    # Extract members of the distribution group "PB_"<Mailbox name>
    $CurrentDistributionGroupMembersArray = Get-DistributionGroupMember -identity ("PB_" + $MailboxString) | Select-Object -ExpandProperty WindowsLiveID # Do not remove, this is for live environment
    #$PB_AchterhoekFACTBegeleiding = Get-DistributionGroupMember -identity ("PB_" + "AchterhoekFACTBegeleiding") | Select-Object -ExpandProperty WindowsLiveID #TEST
    Write-Output "Current Group: $MailboxString" # test
    
    # Get members that currently have access rights
    #$CurrentMailBoxRights = Get-mailbox -Identity achterhoekfactbegeleiding | get-mailboxpermission | Select-Object -ExpandProperty user | where { $_ -like "*@iriszorg.nl"} # TEST
    $CurrentMailBoxRights = Get-mailbox -Identity $Mailbox | get-mailboxpermission | Select-Object -ExpandProperty user | where { $_ -like "*@iriszorg.nl"} -ErrorAction SilentlyContinue # TEST



    Try {    
        If ($CurrentMailBoxRights -eq $Null) {
            Write-Host "There are no mailbox rights set, only adding users.."
            $UsersToAdd = $CurrentDistributionGroupMembersArray
            $UsersToRemove = $Null
        }
        Else {

            # Compare list of users that have mailbox rights with members of the PB_ Security Group 
            $UsersToAdd = compare-object -ReferenceObject $CurrentDistributionGroupMembersArray -DifferenceObject $CurrentMailBoxRights
            Write-Host "Adding following users to [$Mailbox].." -ForegroundColor Green
            $UsersToAdd = $UsersToAdd | where { $_.SideIndicator -like "<="} 
            $UsersToAdd = $UsersToAdd | Select -ExpandProperty InputObject

            $UsersToAdd | Format-Table

            $UsersToRemove = compare-object -ReferenceObject $CurrentMailBoxRights -DifferenceObject $CurrentDistributionGroupMembersArray
            Write-Host "Removing following users to [$Mailbox].." -ForegroundColor Yellow 
            $UsersToRemove = $UsersToRemove | where { $_.SideIndicator -like "<="} 
            $UsersToRemove = $UsersToRemove | Select -ExpandProperty InputObject

            $UsersToRemove | Format-Table
        }

        # Add rights for all users that are new in the PB_ Security Group

        #ForEach ($Member in $CurrentDistributionGroupMembersArray) {
        ForEach ($MemberUser1 in $UsersToAdd) { #TEST
        Write-Output "Current member: $Member1"
                    # Add mailbox permissions with inheritance to child folders within mailbox
            Add-MailboxPermission -Identity $MailboxString -User $MemberUser1 -AccessRights $AccessRights -InheritanceType All -verbose #-whatif
            Add-RecipientPermission -Identity $MailboxString -Trustee $MemberUser1 -AccessRights SendAs -confirm:$False -verbose #-whatif
        }
        

        # Remove rights for all users that currently have rights for the mailbox but aren't member of the PB_ Security group.
        If ($UsersToRemove -ne $Null) { # Don't remove user rights when there are none set!
            ForEach ($MemberUser2 in $UsersToRemove) { 
            Write-Output "Current member: $Member2"
                        # Add mailbox permissions with inheritance to child folders within mailbox
                Remove-MailboxPermission -Identity $MailboxString -User $MemberUser2 -AccessRights $AccessRights -InheritanceType All -Confirm:$false -verbose #-WhatIf
                Remove-RecipientPermission -Identity $MailboxString -Trustee $MemberUser2 -AccessRights SendAs -confirm:$False -verbose #-whatif
            }
        }
        $CurrentMailBoxRights = $Null # reset value 
        $UsersToRemove = $Null # reset value
        # View Result
        Get-MailboxPermission -Identity $Mailbox | Format-Table



    } # End Try
    Catch {
        Write-Output "An error occured with setting rights for [$MailboxString]"

    }


#} # End If
Else {
    continue 
    # NEXT
}


}
Stop-Transcript

























#write-output $CurrentDistributionGroupMembersArray






<#


    $CurrentMemberArray += New-Object -TypeName PSObject -Property @{ # Fill Array with custom objects
        'Group' = $Group
        'members' = $CurrentDistributionGroupMembersArray
    } # End PS Object

    ForEach ($Member in $CurrentDistributionGroupMembersArray) {
        #Write-Output $Member.DisplayName       
        Write-Output $Member
    }
}

write-host $CurrentMemberArray



$NewDistributionGroupMembersArray = @()

# Compare current distribution list members with listed full access rights mailbox
ForEach ($Row in $CSV) {
    
    $MailboxObject = Get-Mailbox $Row.Mailbox
    $MailboxString = $MailboxObject | Select-Object -ExpandProperty userprincipalname
    $DistributionGroup = Get-DistributionGroup -identity $Row.Group 
    $NewDistributionGroupMembersArray = Get-DistributionGroupMember -identity $Row.Group


    # Compare current AD group members with members listed in mailbox rights

    $NewrrentDistributionGroupMembersArray = Get-DistributionGroupMember -identity $Group | Select-Object -ExpandProperty name
    Write-Output "Current Group: $Group"
    
    $NewMemberArray += New-Object -TypeName PSObject -Property @{ # Fill Array with custom objects
        'Group' = $Group
        'members' = $CurrentDistributionGroupMembersArray
    } # End PS Object




    Compare-Object -ReferenceObject $CurrentMemberArray -DifferenceObject $NewDistributionGroupMembersArray

    
    ForEach ($Member in $DistributionGroupMembersArray) {
        Write-Output $Member.DistinguishedName
        
        # Add mailbox permissions with inheritance to child folders within mailbox
        #Add-MailboxPermission -Identity $MailboxString -User $($member.samaccountname) -AccessRights $AccessRights -InheritanceType All  


    }
}






#>





<#
# Loop through CSV and provide access rights:
ForEach ($Row in $CSV) {
    
    $MailboxObject = Get-Mailbox $Row.Mailbox
    $MailboxString = $MailboxObject | Select-Object -ExpandProperty userprincipalname
    #$DistributionGroup = Get-DistributionGroup -identity $Row.Group 
    $NewDistributionGroupMembersArray = @()

    $NewDistributionGroupMembersArray = Get-DistributionGroupMember -identity $Row.Group

    
    ForEach ($Member in $DistributionGroupMembersArray) {
        #Write-Output $Member.DisplayName
        Write-Output $Member.DistinguishedName
        
        # Add mailbox permissions with inheritance to child folders within mailbox
        Add-MailboxPermission -Identity $MailboxString -User $($member.samaccountname) -AccessRights $AccessRights -InheritanceType All  


    }
}
#>








































#>

<#


    
    Write-Log "ALERT: Expand distribution group membership. [$($CheckDelegate.Name)]"
            ForEach ($Member in Get-DistributionGroupMember $CheckDelegate.Name -ResultSize Unlimited) {
                $CheckMember = Get-Recipient $Member -ErrorAction SilentlyContinue
                If ($CheckMember -ne $null) {
                    $DelegateName = $DelegateID + ":" + $CheckMember.Name
                    $DelegateEmail = $CheckMember.PrimarySmtpAddress
                    "$MailboxName,$MailboxEmail,$DelegateName,$DelegateEmail,$DelegateAccess" | Out-File $ExportFile -Append } } }
    
    
    
       
    [string]$MailboxEmail = $Mailbox.PrimarySmtpAddress
    $CheckMailbox = Get-Recipient $MailboxEmail -ErrorAction SilentlyContinue
    If ($CheckMailbox -eq $null) { Write-Log "ERROR: Mailbox not found. [$MailboxEmail]"; Continue }
    [string]$MailboxName = $CheckMailbox.Name
    [string]$MailboxDN = $CheckMailbox.DistinguishedName
    $Progress = $Progress + 1
    Write-Log ""; Write-Log "INFO: Audit mailbox $Progress of $MailboxCount. [$MailboxEmail]"

    # --- Export mailbox access permissions

    If ($IncludeMailboxAccess -eq $true) {
        Write-Log "AUDIT: Mailbox access permissions..."
        $Delegates = @()
        $Delegates = (Get-MailboxPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.IsInherited -eq $false })
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = $Delegate.AccessRights
                Check-Delegates $Delegate.User $MailboxAccessExport } } }

    # --- Export SendAs permissions

    If ($IncludeSendAs -eq $true) {
        Write-Log "AUDIT: Send As permissions..."
        $Delegates = @()
        $Delegates = Get-ADPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.ExtendedRights -like "*send-as*" }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendAs" 
                Check-Delegates $Delegate.User $SendAsExport } } }
















ForEach ($Mailbox in $Mailboxes) {
    [string]$MailboxEmail = $Mailbox.PrimarySmtpAddress
    $CheckMailbox = Get-Recipient $MailboxEmail -ErrorAction SilentlyContinue
    If ($CheckMailbox -eq $null) { Write-Log "ERROR: Mailbox not found. [$MailboxEmail]"; Continue }
    [string]$MailboxName = $CheckMailbox.Name
    [string]$MailboxDN = $CheckMailbox.DistinguishedName
    $Progress = $Progress + 1
    Write-Log ""; Write-Log "INFO: Audit mailbox $Progress of $MailboxCount. [$MailboxEmail]"

    # --- Export mailbox access permissions

    If ($IncludeMailboxAccess -eq $true) {
        Write-Log "AUDIT: Mailbox access permissions..."
        $Delegates = @()
        $Delegates = (Get-MailboxPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.IsInherited -eq $false })
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = $Delegate.AccessRights
                Check-Delegates $Delegate.User $MailboxAccessExport } } }

    # --- Export SendAs permissions

    If ($IncludeSendAs -eq $true) {
        Write-Log "AUDIT: Send As permissions..."
        $Delegates = @()
        $Delegates = Get-ADPermission $MailboxDN | Where { $DelegatesToSkip -notcontains $_.User -and $_.ExtendedRights -like "*send-as*" }
        If ($Delegates -ne $null) {
            ForEach ($Delegate in $Delegates) {
                $DelegateAccess = "SendAs" 
                Check-Delegates $Delegate.User $SendAsExport } } }

                #>
