<#PSScriptInfo

.TYPE Controller

.VERSION 0.1.0

.TEMPLATEVERSION 1

.GUID B1ECB5D8-4E78-44ED-9146-77A678AD3AF6

.AUTHOR Christoph Rust

.CONTRIBUTORS

.COMPANYNAME KrizKodez Tools and Components

.TAGS

.EXTERNALMODULEDEPENDENCIES
    ActiveDirectory, Microsoft Module

.REQUIREDSCRIPTS
 
.EXTERNALSCRIPTDEPENDENCIES

.REQUIREDBINARIES

.DESCRIPTION
    The script allows to create groups with dynamic members from normal groups by adding special parameters
    which implement a rich set of query functionalities.

.RELEASENOTES
    2025-07-22,0.1.0,Christoph Rust,Initial Release

#>

<#
.SYNOPSIS
    Implement AD DS groups with dynamic members.

.DESCRIPTION
    The script allows to create groups with dynamic members from normal groups by adding special parameters
    which implement a rich set of query functionalities.

.INPUTS
    System.String 
    The path of a JSON config file to be used.

    System.String
    An identity value of a domain group.
    
    You cannot pipe input to this function.

.OUTPUTS
    None
    
.NOTES
    The script does not have a rich public interface because the main application
    is the usage with a scheduled task. It supports the -WhatIf switch to test which changes
    it would apply, the command line parameter has priority over the configuration file setting.
    You could submit also a single group with the -Identity parameter.
    
    See the about_UpdateQueryGroup.help textfile for more detailed information.

.EXAMPLE
    .\UpdateQueryGroup.ps1 -Identity TestGroup1

    Updates the dynamic group TestGroup1. The configuration is read from the default configuration file.

.EXAMPLE
    .\UpdateQueryGroup.ps1 -Identity TestGroup1 -WhatIf

    Simulates only the changes in the group TestGroup1.

.EXAMPLE
    .\UpdateQueryGroup.ps1 -ConfigFile C:\Scripts\AD\DynamicGroupConfig\File_Ressource_Groups.json

    Updates dynamic groups defined in a dedicated configuration file.

.PARAMETER Identity
    Specifies an Active Directory group object by providing one of the following property values.
    The identifier in parentheses is the LDAP display name for the attribute. The acceptable values for this parameter are:
    distinguished name, GUID (objectGUID), security identifier (objectSid) or SAM account name (sAMAccountName).

.PARAMETER Configfile
    The path of the JSON config file to be used.
    By default the Config.json is taken from the same directory as the controller.

#>

# PARAMETERS
[CmdletBinding(SupportsShouldProcess)]
param
(
    [String]$Identity,

    [string]$Configfile="$PSScriptRoot\Config.json"
)

# PREREQUISITES
Import-Module -Name ActiveDirectory -ErrorAction Stop

# INCLUDE LIBRARIES
    # PRIVATE
    . "$PSScriptRoot\UpdateQueryGroup.lib.ps1"
    . "$PSScriptRoot\DynamicExpression.UpdateQueryGroup.lib"

    # PUBLIC
    # NA

# DECLARATIONS AND DEFINITIONS
    # ARGUMENTS
    # NA

    # CONSTANTS
    New-Variable -Name SCRIPT_LOG_NAME          -Value 'UpdateQueryGroup.txt' -Option Constant -WhatIf:$false
    New-Variable -Name GROUP_QUERYDATA_LOG_NAME -Value '_QueryParameters'     -Option Constant -WhatIf:$false
    New-Variable -Name REXEX_FAILED_FLAG        -Value '^\s*FAILED\s*\|'      -Option Constant -WhatIf:$false

    # VARIABLES             
    $Domain       = $null                                 # Several Domain data and OU names.
    $FailedGroups = @()                                   # Collect all groups where an error occured.
    $Groups       = $null                                 # All groups to be processed.
    $HostInstance = $null                                 # Computer- and InstanceID name.
    $LogResults   = @()                                   # Collects all log data.
    $Now          = Get-Date -Format 'HH''/''mm''/''ss'   # Script timestamp    
    $Today        = Get-Date -Format 'yyyyMMdd'           # Script timestamp
   

# PARAMETER CHECK
    # Build the Config object from the config file.
    $Content = Get-Content -Path $Configfile -Raw -ErrorAction Stop
    $Config  = ConvertFrom-Json -InputObject $Content -ErrorAction Stop

# CONTROLLER MAIN CODE

# Check the configuration data of the script.
TestConfigurationData -InputObject $Config -ErrorAction Stop

# Set the 'WhatIf' behaviour from the configuration file.
$WhatIfPreference = $Config.WhatIf
# ATTENTION !!! The WhatIf value submitted with the 'WhatIf' script switch will override this setting
# in the PSCmdlet.ShouldProcess method, and only there !!!

# We need various data and names from the domain.
$Domain = GetDomainAndContainerName

# Get the dynamic groups which should be processed.
$Groups = GetDynamicGroup

# We use the 'HostInstance' value in the logs to distinguish different configuration instances or servers.
$HostInstance = "$($Env:COMPUTERNAME)_$($Config.InstanceID)".TrimEnd('_')

# Now we process all dynamic groups.
Write-Host "Found $($Groups.Count) group(s) in domain ($($Domain.Name)) to be checked.`n`rStart processing..." -ForegroundColor Yellow
foreach ($Group in $Groups)
{
    # Collect here all result data regarding the group.
    $GroupLogResults = @()

    # Show some activity which group we are processing.
    Write-Host ("$($Group.Name) - $($Group.Description)".TrimEnd(' - ')) -ForegroundColor Yellow

    # Convert the notes text (info attribute) of the group to the query parameters.
    $GroupQuery = ConvertToQueryData -Note @($Group.info -split '\r\n') -Name $Group.samAccountName

    # Create group log directory. We use the form <ObjectRID>_<samAccountName> for the name to identify renamed groups more easily.
    $GroupDirectory = "$($Group.RIDN)_$($Group.SamAccountName)"
    $null = New-Item -Name $GroupDirectory -Path $Config.LogPath -ItemType Directory -ErrorAction SilentlyContinue -WhatIf:$false

    # We output the query parameters into the group log directory to make troubleshooting a little bit easier.
    [PSCustomObject]$GroupQuery | ConvertTo-Json | Out-File -FilePath "$($Config.LogPath)\$GroupDirectory\$GROUP_QUERYDATA_LOG_NAME.json" -WhatIf:$false

    # If the group is a dynamic group the first info line must be the shebang tag.
    if (-not $GroupQuery.HasShebang)           { continue }

    # If the group should be processed it must be enabled.
    if (-not $GroupQuery.IsEnabled)            { continue }
    
    # Check if the group has been assigned to the same instance like this process.
    if ($GroupQuery.ID -ne $Config.InstanceID) { continue }

    # Skip the group if something is wrong.
    if (-not $GroupQuery.IsValid)
    {
        LogResult -Message "$($Group.Name):Group has not been processed." -Type ERROR
        continue
    }

    # We try to query the dynamic group members and skip the group if it was not successful.
    $QueryResults = GetDynamicGroupMember -Query $GroupQuery
    if (-not $QueryResults.IsValid)
    {
        LogResult -Message "$($Group.Name):Group has not been processed." -Type ERROR
        continue
    }

    # Now we get the current members of the group and devide them into...
    [object[]]$CurrentMembers = Get-ADGroupMember -Identity $Group.samAccountName

    $RightTypeMembers    = @()       # ...all items which have the defined type...
    $WrongTypeMembers    = @()       # ...and all other items.
    foreach ($Member in $CurrentMembers)
    {
        if ($Member.objectClass -eq $GroupQuery.Type) { $RightTypeMembers += $Member.samAccountName }
        elseif ($GroupQuery.ForceType)                { $WrongTypeMembers += $Member.DistinguishedName }
    }

    # We compare the members with the right type with the query data and devide them into...
    $Comparison    = Compare-Object -ReferenceObject $RightTypeMembers -DifferenceObject $QueryResults.SamAccountNames
    
    $AddMembers    = @()       # ...all items which must be added...
    $RemoveMembers = @()       # ...and the items which must be removed from the group.
    foreach ($Item in $Comparison)
    {
        # The indicator "=>" means that the item is only in the new query data and must be added.
        if ($Item.SideIndicator -eq '=>') { $AddMembers    += $Item.InputObject }
        else                              { $RemoveMembers += $Item.InputObject }
    }

    # Count the total amount of changes for this group, without any changes we skip the group.
    $TotalChanges = $AddMembers.Count + $RemoveMembers.Count + $WrongTypeMembers.Count
    Write-Host "`tTotal of changes ($TotalChanges) found for this group." -ForegroundColor Yellow
    if ($TotalChanges -eq 0) { continue }

    # Here we update the group members if the WhatIf mode allows it.
    if (($PSCmdlet.ShouldProcess($Group.Name,"Udate group members")) -and ($GroupQuery.Whatif -eq $false))
    {
        try
        {
            $SharedParameters = @{Identity=$Group.samAccountName; WhatIf=$false; ErrorAction='Stop'}
            if ($AddMembers)       { Add-ADGroupMember    @SharedParameters -Members $AddMembers }
            if ($RemoveMembers)    { Remove-ADGroupMember @SharedParameters -Members $RemoveMembers    -Confirm:$false }
            if ($WrongTypeMembers) { Remove-ADGroupMember @SharedParameters -Members $WrongTypeMembers -Confirm:$false }
            $EventLogType = "INFO"
        }
        catch
        {
            LogResult -Message "$($Group.Name):An exception was thrown ($($_.Exception.Message))." -Type ERROR
            continue
        }
    }
    else { $EventLogType = "WHATIF" }

    # Output information of what we have done so far.
    foreach ($Member in $AddMembers)       { LogResult -Message "Added $($GroupQuery.Type) ($Member)"    -Type $EventLogType -GroupLog }
    foreach ($Member in $RemoveMembers)    { LogResult -Message "Removed $($GroupQuery.Type) ($Member)"  -Type $EventLogType -GroupLog }
    foreach ($Member in $WrongTypeMembers) { LogResult -Message "Removed item with wrong type ($Member)" -Type $EventLogType -GroupLog }

    # We must calculate the effective 'WhatIf' value before calling the next functions,
    # because they are not wrapped into a $PSCmdlet.ShouldProcess methode.
    if ($null -eq $PSBoundParameters.WhatIf) { $EffectiveWhatIf = $Config.WhatIf -or $GroupQuery.WhatIf }
    else                                     { $EffectiveWhatIf = $PSBoundParameters.WhatIf -or $GroupQuery.WhatIf }
    
    # Check if we must push objects to the PushOU.
    $HasPushOUFailed  = $null
    $SharedParameters = @{Scope=$GroupQuery.PushScope; WhatIf=$EffectiveWhatIf; ErrorAction='SilentlyContinue'}
    if ($GroupQuery.PushOU)
    { MoveObjectToContainer @SharedParameters -Identity $AddMembers -Destination $GroupQuery.PushOU -ErrorVariable $HasPushOUFailed }

    # Check if we must pull objects from the PushOU to the PullOU.
    if ($GroupQuery.PullOU -and (-not $HasPushOUFailed))
    { MoveObjectToContainer @SharedParameters -Identity $RemoveMembers -Destination $GroupQuery.PullOU -Source $GroupQuery.PushOU }

    # Flush the log for this group.
    if ($GroupLogResults) { $GroupLogResults | Out-File -FilePath "$($Config.LogPath)\$GroupDirectory\$Today.txt" -Append -WhatIf:$false }

    # Send the mail for this group if necessary.
    if ($GroupQuery.Mail -and $GroupLogResults) { SendMail -Receiver $GroupQuery.Mail -Subject "Updates Dynamic Group: $($Group.Name)"}

}# End of foreach all dynamic groups.

# Update the 'FAILED' flag in the group description.
foreach ($Group in $Groups)
{
    $IsFailedFlagSet = $Group.Description -match $REXEX_FAILED_FLAG
    if ($FailedGroups -contains $Group)
    {
        if ((-not $IsFailedFlagSet) -and $PSCmdlet.ShouldProcess($Group.Name,'Set FAILED flag in group description'))
        { Set-ADGroup -Identity $Group -Description "FAILED | $($Group.Description)" -WhatIf:$false -ErrorAction Stop }
    }
    else
    {
        if ($IsFailedFlagSet -and $PSCmdlet.ShouldProcess($Group.Name,'Clear FAILED flag in group description'))
        {
            $NewDescription = ($Group.Description -replace $REXEX_FAILED_FLAG,'').TrimStart(' ')
            Set-ADGroup -Identity $Group.SamAccountName -Description $NewDescription -WhatIf:$false -ErrorAction Stop    
        }
    }
}# End of foreach all groups updating 'FAILED' flag.

# Write the log if needed.
if ($LogResults)
{ $LogResults | Out-File -FilePath "$($Config.LogPath)\${Today}_${HostInstance}_$SCRIPT_LOG_NAME" -Append -WhatIf:$false }

Write-Host "Finished processing." -ForegroundColor Yellow

# END MAIN CODE

# EXCEPTION HANDLING
trap
{
    "${Now}:ERROR:$($_.Exception.Message)" | Out-File -FilePath "$PSScriptRoot\Error_$Today.txt" -Append -WhatIf:$false
    break
}


