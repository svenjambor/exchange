
  <#
      .SYNOPSIS
      Function to convert/migrate on-premises Exchange distribution group to a Cloud (Exchange Online) distribution group

      .DESCRIPTION
      Copies attributes of a synchronized group to a placeholder group and CSV file. 
      After initial export of group attributes, the on-premises group can have the attribute "AdminDescription" set to "Group_NoSync" which will stop it from be synchronized.
      The "-Finalize" switch can then be used to write the addresses to the new group and convert the name.  The final group will be a cloud group with the same attributes as the previous but with the additional ability of being able to be "self-managed".
      Once the contents of the new group are validated, the on-premises group can be deleted.

      This script should be run on a machine with the ExchangeOnline and AD modules installed (so we can set Group_NoSync). If you do not want his to be set automatically (e.g. because the account running this script has no rights to update the group's AD properties) then use the -doNotSetGroupNoSync switch
      
      The script will try to synchronize the edited Group's attributes to AzureAd (i.e. it will remove the group since we've set "Group_NoSync").  If you do not want this (e.g. because you are batch-editing and want this to happen in a separate step later) then do not set -doNotAutoSyncAfterChanges. In any case, ensure that you DO sync the "Group_NoSync" value has been set and the  changes have been synched BEFORE running -Finalize :-)
      If you DO want to autoSync, make sure -SyncServer is set (easiest is to edit the parameters section below to your environment). Als ensure that the AzureADConnect server is reachbale by remote Powershell from your account and machine.

      Finally, the on-premises group can be hidden if the switch -HideOnPremisesGroup is used in combination with -CreatePlaceHolder

      .PARAMETER Group
      Name of group to recreate.

      .PARAMETER CreatePlaceHolder
      Create placeholder DistributionGroup with a given name.

      .PARAMETER Finalize
      Convert a given placeholder group to final DistributionGroup.

      .PARAMETER HideOnPremisesGroup
      Hides the group from the (on premises) address book after creating the online placeholder (i.e. must be run during  -CreatePlaceholder). This prevents onprem users from using the group.

      .PARAMETER DoNotSetGroupNoSync
      Default behavior of the script is to set the on-premises group's AdminDescription to Group_NoSyc. This stops AzureAD Connect from syncing it, thus removing the group from AzureAD/ExchangeOnline the next time AzureAD Connect runs.  Set this switch if you want to set the attribute manually or have another way of removing the group from sync.
      
      .PARAMETER DoNotAutoSyncAfterChanges
      Default behavior of the script is to invoke Start-ADSyncSyncCycle on the AzureAD Connect server after adding the Group_NoSync attribute and/or changing addresbook availability. sWitch this on to prevent syncing, e.g. if you want to run the script several times and want to sync all changes in one go at the end.

      .PARAMETER SyncServer
      Required if automatic syncing of the changes made by -CreatePlaceHolder is switched on.  Defaults to en empty stirng; you could choose to pre-fill it with your AzureAD Connect server's name to prevent having to set this each runtime

      .PARAMETER ExportDirectory
      Export Directory for internal CSV handling. Defaults to script's location 

      .EXAMPLE
       .\Move-ADGroupToCloud.ps1 -Group "DL-Marketing" -CreatePlaceHolder -autoSyncAfterChanges -syncServer "AzureADConnectServer"

      .EXAMPLE
       .\Move-ADGroupToCloud.ps -Group "DL-Marketing" -Finalize

      .NOTES
      This function is based on the Recreate-DistributionGroup.ps1 script of Joe Palarchio and the Export-DistributionGroup2Cloud.ps1 script by Joerg Hochwald

      License: BSD 3-Clause

      .TODO
      - check for IsDirSynced status - no need to move a group that is cloud only
      - create a switch which allows toggling between hiding the onprem group and creating a onorem contact for its addresses instead
  #>

  param
  (
    [Parameter(Mandatory,
    HelpMessage = 'Name of group to recreate.')]
    [string]  $Group,
    [switch]  $CreatePlaceHolder,
    [switch]  $Finalize,
    [ValidateNotNullOrEmpty()]
    [switch]  $HideOnPremisesGroup,
    [switch]  $DoNotSetGroupNoSync,
    [switch]  $DoNotAutoSyncAfterChanges,
    [Parameter(HelpMessage = 'Make sure you specify a syncserver if you want to sync changes to AzureAD automatically')]
    [string]  $SyncServer = "",
    [string]  $ExportDirectory = $PSScriptRoot
  )

  begin
  {
    # Defaults
    $SCN = 'SilentlyContinue'
    $CNT = 'Continue'
    $STP = 'Stop'

    # Connect & Login to ExchangeOnline (with MFA if required)
    $getSessions = Get-PSSession | Select-Object -Property State, Name
    $isConnected = (@($getSessions) -like '@{State=Opened; Name=ExchangeOnlineInternalSession*').Count -gt 0
    If ($isconnected -eq $false) {
        Connect-ExchangeOnline -prefix o365
    }

    # The switches are "negative" terms but we want booleans to tell us that to do
    if($DoNotSetGroupNoSync.IsPresent){[boolean] $setGroupNoSync = $false} else {[boolean] $setGroupNoSync = $true}
    if($DoNotAutoSyncAfterChanges.IsPresent){[boolean] $autoSyncAfterChanges = $false} else {[boolean] $autoSyncAfterChanges = $true}
    if($HideOnPremisesGroup.IsPresent){[boolean] $hideFromAddressList = $true} else {[boolean] $hideFromAddressList = $false}
  }

  process
  {
    If ($CreatePlaceHolder.IsPresent)
    {
      # Create the Placeholder
      If (((Get-o365DistributionGroup -Identity $Group -ErrorAction $SCN).IsValid) -eq $True)
      {
        # Splat to make it more human readable
        $paramGetDistributionGroup = @{
          Identity      = $Group.Trim()
          ErrorAction   = $STP
          WarningAction = $CNT
        }
        try
        {
          $OldDG = (Get-o365DistributionGroup @paramGetDistributionGroup)
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)

          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        try
        {
          [IO.Path]::GetInvalidFileNameChars() | ForEach-Object -Process {
            $Group = $Group.Replace($_,'_')
          }
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)

          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        $OldName = [string]$OldDG.Name.Trim()
        $OldDisplayName = [string]$OldDG.DisplayName.Trim()
        $OldPrimarySmtpAddress = [string]$OldDG.PrimarySmtpAddress.Trim()
        $OldAlias = [string]$OldDG.Alias.Trim()
    
        
        # Splat to make it more human readable
        $paramGetDistributionGroupMember = @{
          Identity      = $OldDG.Name.Trim()
          ErrorAction   = $STP
          WarningAction = $CNT
        }
        try
        {
          $OldMembers = ((Get-o365DistributionGroupMember @paramGetDistributionGroupMember).Name)
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)

          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        If(!(Test-Path -Path $ExportDirectory -ErrorAction $SCN -WarningAction $CNT))
        {
          Write-Verbose -Message ('  Creating Directory: {0}' -f $ExportDirectory)

          # Splat to make it more human readable
          $paramNewItem = @{
            ItemType      = 'directory'
            Path          = $ExportDirectory
            Force         = $True
            Confirm       = $False
            ErrorAction   = $STP
            WarningAction = $CNT
          }
          try
          {
            $null = (New-Item @paramNewItem)
          }
          catch
          {
            $line = ($_.InvocationInfo.ScriptLineNumber)

            # Dump the Info
            Write-Warning -Message ('Error was in Line {0}' -f $line)

            # Dump the Error catched
            Write-Error -Message $_ -ErrorAction $STP
            break
          }
        }

        # Define variables - mostly for future use
        $ExportDirectoryGroupCsv = $ExportDirectory + '\' + $Group + '.csv'

        try
        {
          # TODO: Refactor in future version
          'EmailAddress' > $ExportDirectoryGroupCsv
          $OldDG.EmailAddresses >> $ExportDirectoryGroupCsv
          'x500:'+$OldDG.LegacyExchangeDN >> $ExportDirectoryGroupCsv
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)

          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        # Define variables - mostly for future use
        $NewDistributionGroupName = 'Cloud-' + $OldName
        $NewDistributionGroupAlias = 'Cloud-' + $OldAlias
        $NewDistributionGroupDisplayName = 'Cloud-' + $OldDisplayName
        $NewDistributionGroupPrimarySmtpAddress = 'Cloud-' + $OldPrimarySmtpAddress

        # TODO: Replace with Write-Verbose in future version of the function
        Write-Output -InputObject ('  Creating Group: {0}' -f $NewDistributionGroupDisplayName)

        # Splat to make it more human readable
        $paramNewDistributionGroup =  @{
          Name               = $NewDistributionGroupName.Trim()
          Alias              = $NewDistributionGroupAlias.Trim()
          DisplayName        = $NewDistributionGroupDisplayName.Trim()
          ManagedBy          = $OldDG.ManagedBy
          Members            = $OldMembers
          PrimarySmtpAddress = $NewDistributionGroupPrimarySmtpAddress
          ErrorAction        = $STP
          WarningAction      = $CNT
        }


        if ($OldDG.GroupType -like "*security*") {
            $paramNewDistributionGroup += @{Type = "security"}
        }

        try
        {
          $null = (New-o365DistributionGroup @paramNewDistributionGroup)
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)
          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        # Wait for 3 seconds
        $null = (Start-Sleep -Seconds 3)

        # Define variables - mostly for future use
        $SetDistributionGroupIdentity = 'Cloud-' + $OldName
        $SetDistributionGroupDisplayName = 'Cloud-' + $OldDisplayName

        # TODO: Replace with Write-Verbose in future version of the function
        Write-Output -InputObject ('  Setting Values For: {0}' -f $SetDistributionGroupDisplayName)

        # Splat to make it more human readable
        $paramSetDistributionGroup = @{
          Identity                               = $SetDistributionGroupIdentity
          AcceptMessagesOnlyFromSendersOrMembers = $OldDG.AcceptMessagesOnlyFromSendersOrMembers
          RejectMessagesFromSendersOrMembers     = $OldDG.RejectMessagesFromSendersOrMembers
          ErrorAction                            = $STP
          WarningAction                          = $CNT
        }
        try
        {
          $null = (Set-o365DistributionGroup @paramSetDistributionGroup)
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)

          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP

          # Something that should never be reached
          break
        }

        # Define variables - mostly for future use
        $SetDistributionGroupIdentity = $('Cloud-' + $OldName.Trim()).Trim()

        # Splat to make it more human readable
        $paramSetDistributionGroup = @{
          Identity                             = $SetDistributionGroupIdentity
          AcceptMessagesOnlyFrom               = $OldDG.AcceptMessagesOnlyFrom
          AcceptMessagesOnlyFromDLMembers      = $OldDG.AcceptMessagesOnlyFromDLMembers
          BypassModerationFromSendersOrMembers = $OldDG.BypassModerationFromSendersOrMembers
          BypassNestedModerationEnabled        = $OldDG.BypassNestedModerationEnabled
          CustomAttribute1                     = $OldDG.CustomAttribute1
          CustomAttribute2                     = $OldDG.CustomAttribute2
          CustomAttribute3                     = $OldDG.CustomAttribute3
          CustomAttribute4                     = $OldDG.CustomAttribute4
          CustomAttribute5                     = $OldDG.CustomAttribute5
          CustomAttribute6                     = $OldDG.CustomAttribute6
          CustomAttribute7                     = $OldDG.CustomAttribute7
          CustomAttribute8                     = $OldDG.CustomAttribute8
          CustomAttribute9                     = $OldDG.CustomAttribute9
          CustomAttribute10                    = $OldDG.CustomAttribute10
          CustomAttribute11                    = $OldDG.CustomAttribute11
          CustomAttribute12                    = $OldDG.CustomAttribute12
          CustomAttribute13                    = $OldDG.CustomAttribute13
          CustomAttribute14                    = $OldDG.CustomAttribute14
          CustomAttribute15                    = $OldDG.CustomAttribute15
          ExtensionCustomAttribute1            = $OldDG.ExtensionCustomAttribute1
          ExtensionCustomAttribute2            = $OldDG.ExtensionCustomAttribute2
          ExtensionCustomAttribute3            = $OldDG.ExtensionCustomAttribute3
          ExtensionCustomAttribute4            = $OldDG.ExtensionCustomAttribute4
          ExtensionCustomAttribute5            = $OldDG.ExtensionCustomAttribute5
          GrantSendOnBehalfTo                  = $OldDG.GrantSendOnBehalfTo
          HiddenFromAddressListsEnabled        = $True
          MailTip                              = $OldDG.MailTip
          MailTipTranslations                  = $OldDG.MailTipTranslations
          MemberDepartRestriction              = $OldDG.MemberDepartRestriction
          MemberJoinRestriction                = $OldDG.MemberJoinRestriction
          ModeratedBy                          = $OldDG.ModeratedBy
          ModerationEnabled                    = $OldDG.ModerationEnabled
          RejectMessagesFrom                   = $OldDG.RejectMessagesFrom
          RejectMessagesFromDLMembers          = $OldDG.RejectMessagesFromDLMembers
          ReportToManagerEnabled               = $OldDG.ReportToManagerEnabled
          ReportToOriginatorEnabled            = $OldDG.ReportToOriginatorEnabled
          RequireSenderAuthenticationEnabled   = $OldDG.RequireSenderAuthenticationEnabled
          SendModerationNotifications          = $OldDG.SendModerationNotifications
          SendOofMessageToOriginatorEnabled    = $OldDG.SendOofMessageToOriginatorEnabled
          BypassSecurityGroupManagerCheck      = $True
          ErrorAction                          = $STP
          WarningAction                        = $CNT
        }
        try
        {
          $null = (Set-o365DistributionGroup @paramSetDistributionGroup)
        }
        catch
        {
          $line = ($_.InvocationInfo.ScriptLineNumber)
          # Dump the Info
          Write-Warning -Message ('Error was in Line {0}' -f $line)

          # Dump the Error catched
          Write-Error -Message $_ -ErrorAction $STP
          break
        }

        # Start AD changes if the -doNotSetNoSync wasn't used 
        If ($setGroupNoSync) {
          If (Get-Module -ListAvailable -Name ActiveDirectory) {
            Write-Verbose "ActiveDirectory Module exists, trying to load"
            try
            {
              Import-Module ActiveDirectory
            }
            catch
            {
              $line = ($_.InvocationInfo.ScriptLineNumber)
              # Dump the Info
              Write-Warning -Message ('Error was in Line {0}' -f $line)
    
              # Dump the Error catched
              Write-Error -Message $_ -ErrorAction $STP
              break
            }
          } 
          else {
              Write-Error "ActiveDirectory Module does NOT exist on this machine. "
              Write-Error "Script will continue WITHOUT setting AD attributes (although you specified otherwise)"
              Write-Error "Make sure you set the attribute manually and run a sync BEFORE doing a -Finalize"

              # No point in running ADConnect sync further along the script of we didn't change anything in AD
              $setGroupNoSync = $false
              $autoSyncAfterChanges = $false
              $hideFromAddressList = $false
          }
        }

        # If we have the module then $setGroupNoSync is still true. Let's set the adminDescription attribute to Group_NoSync (if we can)
        If ($setGroupNoSync){
            try
            {
                $FilterString = '(Name -eq "' + $OldName + '")'
                Get-ADGroup -Filter $FilterString | Set-ADGroup -Replace @{adminDescription='Group_NoSync'}
            }
            catch
            {
                $line = ($_.InvocationInfo.ScriptLineNumber)
                # Dump the Info
                Write-Warning -Message ('Error was in Line {0}' -f $line)
    
                # Dump the Error catched
                Write-Error -Message $_ -ErrorAction $STP
                break
            }
        }
        
        If ($hideFromAddressList) {
          try
          {
              $FilterString = '(Name -eq "' + $OldName + '")'
              Get-ADGroup -Filter $FilterString | Set-ADGroup -Replace @{msExchHideFromAddressLists='TRUE'}
          }
          catch
          {
              $line = ($_.InvocationInfo.ScriptLineNumber)
              # Dump the Info
              Write-Warning -Message ('Error was in Line {0}' -f $line)
  
              # Dump the Error catched
              Write-Error -Message $_ -ErrorAction $STP
              break
          }
        }

        If ($autoSyncAfterChanges -and ($SyncServer -ne "")) {
          try
          {
            write-host "Going to sleep for 60 seconds to give AD a chance to catch up before running AD Connect Sync"
            $null = (Start-Sleep -Seconds 60)
            Invoke-Command -ComputerName $syncServer -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta} -ErrorAction $STP
          }
          catch
          {
            $line = ($_.InvocationInfo.ScriptLineNumber)
            # Dump the Info
            Write-Warning -Message ('Error was in Line {0}' -f $line)
  
            # Dump the Error catched
            Write-Error -Message $_ -ErrorAction $CNT
            break
          }
        }
      }
      Else
      {
        Write-Error -Message ('The distribution group {0} was not found' -f $Group) -ErrorAction $CNT
      }
    }
    ElseIf ($Finalize.IsPresent)
    {
      # Do the final steps

      # Define variables - mostly for future use
      $GetDistributionGroupIdentity = $('Cloud-' + $Group.Trim()).Trim()

      # Splat to make it more human readable
      $paramGetDistributionGroup = @{
        Identity      = $GetDistributionGroupIdentity
        ErrorAction   = $STP
        WarningAction = $CNT
      }
      try
      {
        $TempDG = (Get-o365DistributionGroup @paramGetDistributionGroup)
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)

        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      $TempPrimarySmtpAddress = $TempDG.PrimarySmtpAddress

      try
      {
        [IO.Path]::GetInvalidFileNameChars() | ForEach-Object -Process {
          $Group = $Group.Replace($_,'_')
        }
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)

        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      $OldAddressesPatch = $ExportDirectory + '\' + $Group + '.csv'

      # Splat to make it more human readable
      $paramImportCsv = @{
        Path          = $OldAddressesPatch
        ErrorAction   = $STP
        WarningAction = $CNT
      }
      try
      {
        $OldAddresses = @(Import-Csv @paramImportCsv)
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)

        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      try
      {
        $NewAddresses = $OldAddresses | ForEach-Object -Process {
          $_.EmailAddress.Replace('X500','x500')
        }
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)

        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      $NewDGName = $TempDG.Name.Replace('Cloud-','').Trim()
      $NewDGDisplayName = $TempDG.DisplayName.Replace('Cloud-','').Trim()
      $NewDGAlias = $TempDG.Alias.Replace('Cloud-','').Trim()

      try
      {
        $NewPrimarySmtpAddress = ($NewAddresses | Where-Object -FilterScript {
            $_ -clike 'SMTP:*'
        }).Replace('SMTP:','')
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)
        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      # Splat to make it more human readable
      $paramSetDistributionGroup = @{
        Identity                        = $TempDG.Name.Trim()
        Name                            = $NewDGName.Trim()
        Alias                           = $NewDGAlias.Trim()
        DisplayName                     = $NewDGDisplayName.Trim()
        PrimarySmtpAddress              = $NewPrimarySmtpAddress
        HiddenFromAddressListsEnabled   = $False
        BypassSecurityGroupManagerCheck = $True
        ErrorAction                     = $STP
        WarningAction                   = $CNT
      }
      try
      {
        $null = (Set-o365DistributionGroup @paramSetDistributionGroup)
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)
        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      $paramSetDistributionGroup = @{
        Identity                        = $NewDGName.Trim()
        EmailAddresses                  = @{
          Add = $NewAddresses
        }
        BypassSecurityGroupManagerCheck = $True
        ErrorAction                     = $STP
        WarningAction                   = $CNT
      }
      try
      {
        $null = (Set-o365DistributionGroup @paramSetDistributionGroup)
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)
        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

      # Splat to make it more human readable
      $paramSetDistributionGroup = @{
        Identity                        = $NewDGName.Trim()
        EmailAddresses                  = @{
          Remove = $TempPrimarySmtpAddress
        }
        BypassSecurityGroupManagerCheck = $True
        ErrorAction                     = $STP
        WarningAction                   = $CNT
      }
      try
      {
        $null = (Set-o365DistributionGroup @paramSetDistributionGroup)
      }
      catch
      {
        $line = ($_.InvocationInfo.ScriptLineNumber)

        # Dump the Info
        Write-Warning -Message ('Error was in Line {0}' -f $line)

        # Dump the Error catched
        Write-Error -Message $_ -ErrorAction $STP

        # Something that should never be reached
        break
      }

    }
    Else
    {
      Write-Error -Message "  ERROR: No options selected, please use '-CreatePlaceHolder' or '-Finalize'" -ErrorAction $STP

      # Something that should never be reached
      break
    }
  }
