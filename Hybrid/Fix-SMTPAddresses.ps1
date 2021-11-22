#requires -version 3

<#
.SYNOPSIS
    Script will check which domains exist in Exchange Online and conform mailboxes' address to this list. 

.DESCRIPTION

        The script checks mail addresses of mailboxes found in a csv (csv used to create the migration batch in a later step).

        To do so, it logs into Exchange Online to retrieve all accepted domains and compares the mailboxes' addresses to this list.

        Addresses which are incompatible are removed from the on-premise account.

        If switch autoSyncAfterChanges is set and is syncServer filled then the changes are synchronized to Microsoft 365 automatically.

        You can use the Param() section to set default values for your project (such as the sync server name etc).

        The script assumes that it's run with onprem Exchange Powershell loaded and Exchange Online PS V2 present.

.PARAMETER InputCSV
    CSV with mailboxes; scripts expects a column "EmailAddress" to be preset (in other words use the CSV used to create the batch)

.PARAMETER Delimiter
    Which character to use as delimiter; defaults to ",". Use ";" if you a CSV straight from Excel

.PARAMETER autoSyncAfterChanges
    If your account is allowed to start an AzureAD Connect sync, thenset this to $true if you want to sync the objects after fixing them

.PARAMETER syncServer
    Specific batch to query; script will default to all "synced" queries if omitted

.PARAMETER workingDirectory
    Defaults to the script's working directory

.INPUTS
    none

.OUTPUTS
  Script will create CSV file in the current location named skippedItemsList.csv (or skippeditemList [BatchName].csv if a specific batch was queried)

.NOTES
    Version:        1.0
    Author:         Sven Jambor
    Creation Date:  15-03-2019

  
.EXAMPLE
    .\Fix-SMTPAddresses.ps1 -InputCSV '.\BatchCSVFiles\20190315 - Batch 1.csv'
#>

[CmdletBinding()]
Param (
    [String]  $InputCSV,
    [String]  $Delimiter = ',',
    [Boolean] $autoSyncAfterChanges = $true,
    [String]  $syncServer = "AADConnectServer",
    [String]  $workingDirectory =  $PSScriptRoot
    )

Begin {
    #Connect & Login to ExchangeOnline (MFA)
    $getSessions = Get-PSSession | Select-Object -Property State, Name
    $isConnected = (@($getSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
    If ($isconnected -eq $false) {
        Connect-ExchangeOnline -prefix o365
    }

    #------ check input file ------------------------
    #check / fix input file name if it include a path
    $splitArray = @()
    $splitArray = $InputCSV -Split('\\')
    #-- could split the filename with slashes, this means it's a path
    if($splitArray.count -gt 1) {
        $csvFileName = $splitArray[$splitArray.count - 1]
        $csvPath = $InputCSV.Replace($csvFileName,"")
        }
    else {
        #-- no slashes in the filename, so assuming we need a path
        $CsvFileName = $InputCSV
        $csvPath = $workingDirectory + '\'
        }
        
    $csvFullPath = "$csvPath$CsvFileName"
    #------ 

    #---- import CSV files ----------------------------
   $users = $(Import-Csv -Path "$csvFullPath" -Delimiter "$Delimiter").EmailAddress

    #Get accepted (online) domains deduce MS mail domain
    $exoDomains = (Get-o365AcceptedDomain).DomainName | sort
    $msMailSuffix = $exoDomains | Where-Object {$_ -like "*mail.onmicrosoft.com"}


}

Process{
    $bnChangesMade = $false
    if($exoDomains -ne $null -or $exoDomains -ne ""){
        foreach($user in $users){

            $mbx = get-mailbox "$user"
    

            if($mbx -ne $null){

                #extra check: check for ms.mail address based on UPN and add if needed
                $msMailAddress ="$($($mbx.UserPrincipalName).Split('@')[0])@$msMailSuffix"
                $smtpAddresses = ($mbx.EmailAddresses | where {$_.PrefixString -eq "smtp"}).SmtpAddress

                if($smtpAddresses -notcontains $msMailAddress){
                    write-host "setting msMail address for $user" -ForegroundColor Cyan
                    Set-Mailbox "$user" -EmailAddresses @{add=$msMailAddress}
                    $bnChangesMade = $true
                }


                # Cycle through mailbox' smtp addresses & compare domain-portion of string to accepted domains.
                # Remove the address if the domain does not exist in O365
         
                foreach($smtpAddress in $smtpAddresses){
                    $smtpDomain = $smtpAddress.split("@")[1]
                    if($exoDomains -notcontains $smtpDomain){
                        write-host "removing address $smtpAddress" -ForegroundColor Yellow
                        Set-Mailbox "$user" -EmailAddresses @{remove=$smtpAddress}
                        $bnChangesMade = $true
                    }  
                }   

            } else {write-error "No mailbox for $user ?"}
    
            #$mbx.PrimarySmtpAddress.Address
        }


        #start ADConnect sync if needed
        if($autoSyncAfterChanges -and $bnChangesMade){
            write-host "Syncing to AzureAD... pls wait 60 seconds" -ForegroundColor White
            sleep 60
            Invoke-Command -ComputerName $syncServer -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}
        }
    }
}