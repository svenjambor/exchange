#requires -version 3

<#
.SYNOPSIS
    Creates a CSV report for all failed items in a (single or set of) migration batch(es) for mailboxes having "investigate" status or worse.

.DESCRIPTION
    

    The script can query a single batch specified by -BatchName; if this is omitted, all current Synced jobs will be queried.

    Output is a csv file in the current directory containing as much information about the skipped items as possible.

    The CSV can then be used to notify users about items which will not be migrated to Exchange Online.


.PARAMETER BatchName
    Specific batch to query; script will default to all "synced" queries if omitted


.INPUTS
    none

.OUTPUTS
  Script will create CSV file in the current location named skippedItemsList.csv (or skippeditemList [BatchName].csv if a specific batch was queried)

.NOTES
    Version:        1.0
    Author:         Sven Jambor
    Creation Date:  15-11-2021

  
.EXAMPLE
    .\Report-ConsolidatedSkippedItems.ps1 -BatchName "20211115*"
#>


[CmdletBinding()]
Param (
    [String] $BatchName
    )

Begin{

    #Connect & Login to ExchangeOnline (MFA)
    $getSessions = Get-PSSession | Select-Object -Property State, Name
    $isConnected = (@($getSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
    If ($isConnected -eq $false) {
        Connect-ExchangeOnline
    }

    
    #Get batches to deal with
    if($BatchName -eq ""){
        "no such batch"
        $batches = Get-MigrationBatch -Status Synced
    } else {
        Try{
            $batches = Get-MigrationBatch -ErrorAction Stop | Where-Object {$_.Identity.Name -like $BatchName}
            $filenameAppendix = " $($BatchName)".Replace("*","").Replace(" ","")
        } Catch {
            write-error "Could not retrieve batch $($BatchName)"
            Break
        }
    }

    #Set up variable for results
    $skippedItemsList =  New-Object -TypeName 'System.Collections.ArrayList'

}

Process{

    foreach($batch in $batches){
    
        $migrationusers = Get-MigrationUser -BatchId $batch.Identity.Name
        foreach ($mbx in $migrationusers) {
            if($mbx.DataConsistencyScore -eq "Investigate" -or $mbx.DataConsistencyScore -eq "Poor"){
                $mbstats = Get-MigrationUserStatistics -Identity $mbx.MailboxIdentifier -IncludeSkippedItems -IncludeReport 
                $skippedMessages = $mbstats.SkippedItems | Where-Object {$_.Kind -ne "CorruptFolderACL" -and $_.Kind -ne "CorruptMailboxSecurityDescriptor" -and $_.Kind -ne "CorruptFolderRule"} | `
                    select BatchName,Mailbox,Kind,FolderName,Sender,Recipient,Subject,@{name="MessageSize"; expression={[math]::Round($_.MessageSize/1Mb,2)}},DateSent,DateReceived

                foreach($skippedMessage in $skippedMessages){
                    $skippedMessage.BatchName = $batch.Identity.Name
                    $skippedMessage.Mailbox = $mbx.MailboxIdentifier
                    $count = $skippedItemsList.Add($skippedMessage)
            
                }
            }
        }
    }
    
    $skippedItemsList | Export-csv -Path ".\skippedItemsList$($filenameAppendix).csv" -Delimiter ";" -NoTypeInformation

}
