<#PSScriptInfo

.TYPE Private Library

.TEMPLATEVERSION 1

.GUID AA1BEDF2-1E0D-4DE5-A070-A38E6653E81B

.FUNCTIONS
    ConvertToQueryData
    ExcludeObject
    GetDomainAndContainerName
    GetDynamicGroup
    GetDynamicGroupMember
    LogResult
    TestConfigurationData
    TransformSamAccountName

.AUTHOR Christoph Rust

.CONTRIBUTORS

.COMPANYNAME KrizKodez Tools and Components

.TAGS

.EXTERNALMODULEDEPENDENCIES

.REQUIREDSCRIPTS
    UpdateQueryGroup.ps1
 
.EXTERNALSCRIPTDEPENDENCIES
 
.REQUIREDBINARIES

.RELEASENOTES 
    2025-07-22,All functiuons,Christoph Rust,Initial Release

.Description
    This library contains private functions for the scripts defined in REQUIREDSCRIPTS.
#>

# DECLARATIONS AND DEFINITIONS
    # CONSTANTS
    New-Variable -Name VALID_MAILADDRESS_REGEX       -Value '^[a-zA-Z0-9._-]+@([a-zA-Z0-9.-]+\.[a-zA-Z]{2,})$' -Option Constant -WhatIf:$false
    New-Variable -Name VALID_SHORT_MAILADDRESS_REGEX -Value '^[a-zA-Z0-9._-]+$'                                -Option Constant -WhatIf:$false
    New-Variable -Name VALID_INSTANCEID_REGEX        -Value '^[a-z0-9_]+$'                                     -Option Constant -WhatIf:$false
    New-Variable -Name VALID_EXCLUDE_PARAMETER_REGEX -Value '(^[a-z][a-z0-9]*)\[(.*?)\]$'                      -Option Constant -Whatif:$false
    New-Variable -Name VALID_SHEBANG_REGEX           -Value '' -Option Constant -Whatif:$false                                     

    # VARIABLES
    # NA

# SCRIPTBLOCKS
# NA


# FUNCTIONS
function ConvertToQueryData
{
<#
.DESCRIPTION
    Convert the group 'info' attribute text to the query parameter data.
    
.INPUTS
    System.String
    The text data of the info attribute.
    
    You cannot pipe input to this function.

.OUTPUTS
    System.Collections.Hashtable
    The converted query data and status codes. 

.PARAMETER Note
    Array with the text data of the LDAP 'info' attribute of a group.
    
.PARAMETER Name
    Give the query data the name of the group.
#>

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [AllowEmptyString()]   
        [string[]]$Note,

        [string]$Name
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $AllowedQueryParameters = @('Type',
                                    'ForceType',
                                    'WhatIf',
                                    'Filter',       
                                    'LdapFilter',
                                    'SearchBase',
                                    'SearchScope',
                                    'Where',
                                    'VMember',
                                    'Exclude',
                                    'PushOU',
                                    'PushScope',
                                    'PullOU',
                                    'SamIn',
                                    'SamOut',
                                    'ID',
                                    'Mail'
                                   )  
                $Result = [ordered]@{
                                    'Name'           = $Name
                                    'ID'             = $null
                                    'IsValid'        = $true
                                    'IsEnabled'      = $true
                                    'WhatIf'         = $false
                                    'HasRecycleBin'  = $true
                                    'Type'           = $null
                                    'ForceType'      = $false
                                    'Mail'           = @()
                                    'Attributes'     = @('msDS-parentdistname')
                                    'VMembers'       = @()
                                    'Filter'         = $null
                                    'LDAPFilter'     = $null
                                    'SearchBases'    = @()
                                    'SearchScope'    = 'OneLevel'
                                    'Where'          = $null
                                    'HasScriptblock' = $false
                                    'Excludes'       = @{}
                                    'HasExcludes'    = $false
                                    'PushOU'         = $null
                                    'PushScope'      = 'Subtree'
                                    'PullOU'         = $null
                                    'SamIn'          = $null
                                    'SamOut'         = $null
                                    'HasShebang'     = $true
                                    }

    # FUNCTION MAIN CODE

    # Check the first two lines because they are mandatory.
    # The first line must be the shebang tag.
    if ($Note[0] -ne '#!QueryGroup')
    {
        $Result.HasShebang = $false
        return
    }

    # The second line must be a correct formatted 'Enabled' parameter.
    if ($Note[1] -match '^Enabled:(true|false)$')
    { if($matches[1] -eq 'false') { $Result.IsEnabled = $false } }
    else
    {
        LogResult -Message "$($Group.Name):Invalid value of parameter (Enabled)." -Type ERROR
        $Result.IsValid   = $false
        $Result.IsEnabled = $false
        return
    }

    # Analyze the rest of the notes text data.                
    for ($i = 2; $i -lt $Note.Count; $i++)
    {
        # Check if the query could use the 'RecycleBin' token.
        if ((-not $Domain.IsRecycleBinEnabled) -and (-not $Config.ForceRecycleBin))
        { $Result.HasRecycleBin = $false }

        # Check if line is empty.
        if (-not $Note[$i]) { continue }
        
        # Note parameter must have the following format [ParameterName]:[ParameterValue].
        if ($Note[$i].Trim(' ') -notmatch '^([a-z]+):(.*)') 
        {
            $Result.IsValid     = $false
            LogResult -Message "$($Group.Name):$i Invalid parameter format ($($Note[$i]))." -Type ERROR
            continue
        }

        # These are the parameter name and value read from the group 'info' attribute...
        $ParameterNameRaw  = $Matches[1].Trim(' ')
        $ParameterValue    = $Matches[2].Trim(' ')
        
        # ... but because we are using an OrderedDictionary for the $Result variable the Dictionary-Keys are case-sensitive
        # and we must convert the raw parameter name into the correct spelling.
        $ParameterName = $AllowedQueryParameters -match "^${ParameterNameRaw}$"
        
        # Skip unknow parameter names.
        if(-not $ParameterName)
        {
            $Result.IsValid = $false
            LogResult -Message "$($Group.Name):Invalid parameter ($($ParameterNameRaw))." -Type ERROR
            continue
        }

        # Skip empty parameter.
        if(-not $ParameterValue) { continue }

        # Assign the parameters to the 'Result' hashtable.
        switch ($ParameterName)
        {
            { @('WhatIf','ForceType') -contains $ParameterName }
            {
                if(($ParameterValue -eq 'true') -or ($ParameterValue -eq 'false')) { $Result."$ParameterName" = [System.Convert]::ToBoolean($ParameterValue) }
                else
                {
                    $Result."$ParameterName" = $ParameterValue
                    $Result.IsValid          = $false
                    LogResult -Message "$($Group.Name):Parameter ($ParameterName) has invalid value ($ParameterValue)." -Type ERROR
                }
                break
            }

            'Mail'
            {
                switch -Regex ($ParameterValue)
                {
                    "$VALID_MAILADDRESS_REGEX"
                    {
                        $MailDomainName = $Matches[1]
                        if ($Config.Mail.Domains -contains $MailDomainName) { $Result.Mail += $ParameterValue }
                        else
                        {
                            $Result.IsValid = $false
                            LogResult -Message "$($Group.Name):Parameter (Mail) has forbidden mail domain ($ParameterValue)." -Type ERROR
                        }
                    }

                    "$VALID_SHORT_MAILADDRESS_REGEX"
                    { $Result.Mail += "$ParameterValue@$($Config.Mail.Domains[0])" }

                    Default
                    {
                        $Result.Mail   += $ParameterValue
                        $Result.IsValid = $false
                        LogResult -Message "$($Group.Name):Parameter (Mail) has invalid mail address ($ParameterValue)." -Type ERROR
                    }
                }
                break
            }

            'ID'
            {
                $Result.ID = $ParameterValue
                # The ID could consist of alphanumeric characters and the underline.
                if ($ParameterValue -notmatch $VALID_INSTANCEID_REGEX)
                {
                    $Result.IsValid          = $false
                    LogResult -Message "$($Group.Name):Parameter (ID) has invalid value ($ParameterValue)." -Type ERROR
                }
                break
            }
            
            'LDAPFilter'
            {
                $Result.LDAPFilter = $ParameterValue
                break
            }

            'Filter'
            {
                # First we search for dynamic expressions, invoke them and inject the result into the 'Filter' parameter.
                foreach ($Expression in $FunctionOfDynamicExpression.Keys)
                {
                    if ($ParameterValue -notmatch "${Expression}\[(.*?)\]") { continue }
                    
                    $FunctionName       = $FunctionOfDynamicExpression[$Expression]
                    $FunctionParameters = $Matches[1]
                    try { $Expression = Invoke-Expression -Command "$FunctionName $FunctionParameters" }
                    catch
                    {
                        $Result.IsValid = $false
                        LogResult -Message "$($Group.Name):Dynamic Expression in parameter (Filter) returns an exception ($($_.Exception.Message))." -Type ERROR
                        break # The foreach.
                    }

                    $ParameterValue = $ParameterValue.Replace("$FunctionName[$FunctionParameters]",$Expression)

                }# End of foreach Dynamic Expressions.
                
                # We search for attributes or aliases.
                $AttributeMatches = (Select-String -InputObject $ParameterValue -Pattern '[( ]*\s*([a-z0-9-]{2,})\s+-[a-z]{2}\s*' -AllMatches).Matches
                foreach ($AttributeMatch in $AttributeMatches)
                {
                    $FoundAttributeName = $AttributeMatch.Groups[1].Value
                    if ($Config.LDAPAliases."$FoundAttributeName")
                    {
                        $LDAPAttributeName = $Config.LDAPAliases."$FoundAttributeName"
                        $ParameterValue    = $ParameterValue.Replace("$FoundAttributeName","$LDAPAttributeName")
                    }
                    else { $LDAPAttributeName = $FoundAttributeName }
                    
                    # Add the attributes found to the 'Attributes' array to ensure the query will load it.
                    if ($Result.Attributes -notcontains $LDAPAttributeName) { $Result.Attributes += $LDAPAttributeName }
                }

                $Result.Filter = $ParameterValue
                break
            }# End of switch value 'Filter'.

            'Type'
            {
                $Result.Type = $ParameterValue
                if (@('User','Computer') -notcontains $ParameterValue)
                {
                    $Result.IsValid     = $false
                    LogResult -Message "$($Group.Name):Parameter (Type) has invalid value ($ParameterValue)." -Type ERROR
                    break
                }
            }

            { @('SearchScope','PushScope') -contains $ParameterName }
            {
                $Result."$ParameterName" = $ParameterValue
                if (@('OneLevel','Subtree') -notcontains $ParameterValue)
                {
                    $Result.IsValid     = $false
                    LogResult -Message "$($Group.Name):Parameter ($ParameterName) has invalid value ($ParameterValue)." -Type ERROR
                    break
                }
            }

            'VMember'
            {
                $Result.VMembers += $ParameterValue
                break
            }

            { @('SamIn','SamOut') -contains $ParameterName }
            {
                $Result."$ParameterName" = $ParameterValue
                break
            }

            { @('SearchBase','PullOU','PushOU') -contains $ParameterName }
            {
                $FoundDistinguishedName = $ParameterValue
                $ErrorMessage           = $null

                # We try to find the DistinguishedName with different methods.
                switch -regex ($ParameterValue)
                {
                    # OU distinguishedName search.
                    '(OU=|CN=|DC=)'
                    {
                        $FoundDNs = @($Domain.OUDistinguishedNames -match "^$ParameterValue")
                        if ($FoundDNs.Count -eq 1) { $FoundDistinguishedName = $FoundDNs[0] }
                        else                       { $ErrorMessage = "$($Group.Name):Parameter ($ParameterName) contains unknown/ambiguous OU value ($ParameterValue)." }
                    }

                    # OU canonicalName search.
                    '/'
                    {
                        $ParameterValue = $ParameterValue -replace '^/',"$($Domain.DNSName)/"
                        
                        # We have to differentiate between a canonicalName starts with the domain name...  
                        if ($ParameterValue.StartsWith($Domain.DNSName))
                        {
                            if ($Domain.OUCanonicalNames -contains $ParameterValue)
                            { $FoundDistinguishedName = $Domain.DNofCanonicalName[$ParameterValue] }
                            else                                                    
                            { $ErrorMessage = "$($Group.Name):Parameter ($ParameterName) contains unknown OU value ($ParameterValue)." }
                        }
                        # ...or ends with the domain name.
                        # In this case we allow to submit a canonicalName as short as possible but we must check if it is unambiguous.
                        else
                        {
                            $FoundCNs = @($Domain.OUReverseCanonicalNames -match "^$ParameterValue")
                            if ($FoundCNs.Count -eq 1) { $FoundDistinguishedName = $Domain.DNofCanonicalName[$FoundCNs[0]] }
                            else                       { $ErrorMessage = "$($Group.Name):Parameter ($ParameterName) contains unknown/ambiguous OU value ($ParameterValue)."}    
                        }
                    }

                    # OU Name attribute search.
                    default
                    {
                        if ($Domain.AmbiguousOUNames -notcontains $ParameterValue)
                        {
                            if ($Domain.DNofOUName.ContainsKey($ParameterValue))
                            { $FoundDistinguishedName = $Domain.DNofOUName[$ParameterValue] }
                            else                                                 
                            { $ErrorMessage = "$($Group.Name):Parameter ($ParameterName) contains unknow OU name ($ParameterValue)." }
                        }
                        else { $ErrorMessage = "$($Group.Name):Parameter ($ParameterName) contains ambiguous OU name value ($ParameterValue)."}
                    }

                }# End of switch between different OU patterns.
                
                # Check if SearchBase is the Domain Root, which is not allowed if we have disabled the root search.
                if (($ParameterName -eq 'SearchBase') -and ($FoundDistinguishedName -eq $Domain.DistinguishedName) -and (-not $Config.IsRootSearchAllowed))
                {
                    $Result.IsValid = $false
                    $ErrorMessage   = "$($Group.Name):Domain Root in parameter (SearchBase) is not allowed."
                }

                # Check if PushOU or PullOU is the Domain Root, which is not allowed.
                if ((@('PushOU','PullOU') -contains $ParameterName) -and ($FoundDistinguishedName -eq $Domain.DistinguishedName))
                {
                    $Result.IsValid = $false
                    $ErrorMessage   ="$($Group.Name):Domain Root in parameter ($ParameterName) is not allowed."
                }

                # Check if PushOU or PullOU is the Recycle Bin but the 'Recycle Bin' feature is not enabled.
                if ((@('PullOU','PushOU') -contains $ParameterName) -and ($FoundDistinguishedName -eq 'Recycle Bin') -and (-not $Result.HasRecycleBin)) 
                {
                    $Result.IsValid = $false
                    $ErrorMessage   = "$($Group.Name):Parameter ($ParameterName) contains the Recycle Bin but this feature is disabled."
                }

                if ($ErrorMessage)
                {
                    $Result.IsValid         = $false
                    LogResult -Message $ErrorMessage -Type ERROR
                }
                
                if ($ParameterName -eq 'SearchBase') { $Result.SearchBases     += $FoundDistinguishedName }
                else                                 { $Result."$ParameterName" = $FoundDistinguishedName }
           
            }# End of switch value 'SearchBase | PushOU |PullOU'.

            'Exclude'
            {
                $Result.Exclude = $ParameterValue
                
                # In the first capture group we have the LDAP attribute name, in the second capture group is the RexEx for the exclusion operation.
                if ($ParameterValue -notmatch $VALID_EXCLUDE_PARAMETER_REGEX)
                {
                    $Result.IsValid     = $false
                    LogResult -Message "$($Group.Name):Parameter (Exclude) has invalid value ($($ParameterValue))." -Type ERROR
                    break
                }
                
                # Here we check if an alias for the LDAP attribute name has been used and if so we replace it.
                if ($Config.LDAPAliases."$($Matches[1])") { $LDAPAttributeName = $Config.LDAPAliases."$($Matches[1])" }
                else                                      { $LDAPAttributeName = $Matches[1] }
                
                $Result.Excludes.Add($LDAPAttributeName,$Matches[2])
                $Result.HasExcludes = $true

                # Add the attributes found to the 'Attributes' array to ensure the query will load it.
                if ($Result.Attributes -notcontains $LDAPAttributeName) { $Result.Attributes += $LDAPAttributeName }
            }

            'Where'
            {
                # Check if a { or } is missing in the scriptblock definition.
                if ($ParameterValue -match '^\{.*[^}]$' -or $ParameterValue -match '^[^{].*\}$')
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):Parameter (Where) has invalid scriptblock." -Type ERROR
                    break
                }

                # Check if we have a scriptblock {...}.
                if($ParameterValue -match '^\{.*?\}$')
                {
                    $Result.HasScriptblock = $true
                    
                    # We must catch all attribute names which are used in the scriptblock like ?_.Attribute and
                    # check if it is a configured LDAP alias and replace this alias with the correct LDAP attribute name.
                    $AttributeMatches = (Select-String -InputObject $ParameterValue -Pattern '\?_\.([a-z0-9-]+)' -AllMatches).Matches
                    foreach ($AttributeMatch in $AttributeMatches)
                    {
                        $FoundAttributeName = $AttributeMatch.Groups[1].Value
                        if ($Config.LDAPAliases."$FoundAttributeName")
                        {
                            $LDAPAttributeName = $Config.LDAPAliases."$FoundAttributeName"
                            $ParameterValue    = $ParameterValue.Replace("?_.$FoundAttributeName","?_.$LDAPAttributeName")
                        }
                        else { $LDAPAttributeName = $FoundAttributeName }
                        
                        # Add the attributes found to the 'Attributes' array to ensure the query will load it.
                        if ($Result.Attributes -notcontains $LDAPAttributeName) { $Result.Attributes += $LDAPAttributeName }

                    }
                    $Result.Where = $ParameterValue
                    break
                }

                # Here we expect the PowerShell Simplified Sytax for the Where cmdlet.
                switch -RegEx ($ParameterValue)
                {
                    '^-Property\s+([a-z0-9-]+) (.*)$'
                    {   
                        $HasAttributeFound  = $true
                        $NewParameterValue  = "-Property !!! $($Matches[2])"
                        $FoundAttributeName = $Matches[1]
                    }

                    '^([a-z0-9-]+)\s+(-[a-z]{2}.*)$'
                    {
                        $HasAttributeFound  = $true
                        $NewParameterValue  = "!!! $($Matches[2])"
                        $FoundAttributeName = $Matches[1]
                    }

                    '^-Not\s+"*([a-z0-9-]+)"*' # PowerShell 7
                    {
                        if ($PSVersionTable.PSVersion.Major -eq 7)
                        {
                            $HasAttributeFound  = $true
                            $NewParameterValue  = "-Not !!!"
                            $FoundAttributeName = $Matches[1]
                        }
                        else
                        {
                            $HasAttributeFound  = $false
                            LogResult -Message "$($Group.Name):In parameter (Where) the -NOT operator is only existing in PowerShell 7." -Type ERROR  
                        }
                    }

                    '^(.*?) -NotIn -Property ([a-z0-9-]+)$'
                    {
                        $HasAttributeFound  = $true
                        $NewParameterValue  = "$($Matches[1]) -NotIn -Property !!!"
                        $FoundAttributeName = $Matches[2]
                    }

                    default { $HasAttributeFound = $false }
                }

                # We must check if the attribute is defined through its LDAPAlias.
                if ($HasAttributeFound)
                {
                    if ($Config.LDAPAliases."$FoundAttributeName")
                    {
                        $LDAPAttributeName = $Config.LDAPAliases."$FoundAttributeName"
                        $ParameterValue    = $NewParameterValue.Replace('!!!',$LDAPAttributeName)
                    }
                    else { $LDAPAttributeName = $FoundAttributeName }

                    # Add the attributes found to the 'Attributes' array to ensure the query will load it.
                    if ($Result.Attributes -notcontains $LDAPAttributeName) { $Result.Attributes += $LDAPAttributeName }
                    
                    $Result.Where = $ParameterValue
                }
                else
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):Could not found an attribute name in parameter (Where)." -Type ERROR
                }
            }# End of Where parameter.

        }# End of switch between allowed parameters.
    }# End of analyzing notes text.

    # Check that 'Type' has been defined.
    if (-not $Result.Type)
    {
        $Result.IsValid = $false
        LogResult -Message "$($Group.Name):Group has no object Type defined." -Type ERROR
    }

    # Check if a scriptblock is allowed in the 'Where' parameter.
    if ($Result.HasScriptblock -and (-not $Config.IsScriptblockAllowed))
    {
        $Result.IsValid = $false
        LogResult -Message "ERROR:$($Group.Name):Usage of scriptblocks has been disabled." -Type ERROR
    }

    # Check that we only have one filter.
    if ($Result.Filter -and $Result.LDAPFilter)
    {
        $Result.IsValid = $false
        LogResult -Message "ERROR:$($Group.Name):Group has mutually exclusive filter defined." -Type ERROR
    }

    # If we do not have any filter we define a default.
    if ((-not $Result.Filter) -and (-not $Result.LDAPFilter)) { $Result.Filter = '*'}

    # We must handle the case neither SearchBases nor VMembers are defined.
    if((-not $Result.SearchBases) -and (-not $Result.VMembers))
    {
        if ($Config.IsRootSearchAllowed)
        {
            $Result.SearchBases += $Domain.DistinguishedName
            $Result.SearchScope = 'Subtree'
        }
        else
        {
            $Result.IsValid = $false
            LogResult -Message "$($Group.Name):Group has neither SearchBases nor VMembers defined." -Type ERROR
        }
    }

    # Here we check some PushOU-PullOU combinations that must be filtered out.
    if ($Result.PushOU -and $Result.PullOU)
    {
        # A PullOU with a PushOU that is 'Recycle Bin' does not make sense.
        if ($Result.PushOU -eq 'Recycle Bin')
        {
            $Result.IsValid = $false
            LogResult -Message "$($Group.Name):PullOU is not needed if the PushOU is the Recycle Bin." -Type ERROR
        }

        # Check here that the PullOU does not point to the PushOU scope.
        switch ($Result.PushScope)
        {
            'OneLevel'
            { 
                if($Result.PushOU -eq $Result.PullOU)
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):PullOU cannot point to PushOU." -Type ERROR  
                }
            }
            'Subtree'
            {
                if ($Result.PullOU.EndsWith($Result.PushOU))
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):PullOU cannot point to PushOU or a sub OU." -Type ERROR 
                }
            }
        }# End of switch the PushScope.
    }
    else
    {
        # A PullOU without a PushOU is not allowed.
        if ($Result.PullOU)
        {
            $Result.IsValid = $false
            LogResult -Message "$($Group.Name):PullOU without PushOU is not allowed." -Type ERROR 
        }
    }

    # Here we check some SearchBase-PullOU combinations that must be filtered out.
    if ($Result.SearchOU -and $Result.PullOU)
    {
        # Check here that the PullOU does not point to the SearchBase scope.
        switch ($Result.SearchScope)
        {
            'OneLevel'
            { 
                if($Result.SearchScope -eq $Result.PullOU)
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):PullOU cannot point to SearchBase." -Type ERROR    
                }
            }
            'Subtree'
            {
                if ($Result.PullOU.EndsWith($Result.SearchBase))
                {
                    $Result.IsValid = $false
                    LogResult -Message "$($Group.Name):PullOU cannot point to SearchBase or a sub OU." -Type ERROR  
                }
            }

        }# End of switch the SearchScope.
    }

    # Here we check some SamIn-SamOut combinations that must be filtered out.
    if ($Result.SamIn -and (-not $Result.SamOut))
    {
        $Result.IsValid = $false
        LogResult -Message "$($Group.Name):The SamIn regular expression needs a SamOut processing." -Type ERROR 
    }

    if ((-not $Result.SamIn) -and $Result.SamOut)
    {
        $Result.IsValid = $false
        LogResult -Message "$($Group.Name):The SamOut parameter does not work without a SamIn regular expression." -Type ERROR  
    }

    Write-Output $Result

}# End of function ConvertToQueryData

function ExcludeObject
{
<#
.SYNOPSIS
    Exclude objects from a dynamic group.

.DESCRIPTION
    The function applies the 'Exclude' rules and also a Where-Object command
    defined in the dynamic group parameters.    
    
.INPUTS
    System.Object
    User or Computer AD objects.

    System.Collections.Hashtable
    Exclude rules and Where-Object command arguments.

    You cannot pipe input to this function.

.OUTPUTS
    System.Object
    List of AD objects that were not filtered out.

.PARAMETER InputObject
    List of AD objects that need to be checked.

.PARAMETER Exclude
    This paramter contains a hashtable with rules to exclude objects.
    The Key is the LDAP attribute name and the value is a regex

.PARAMETER Where
    

#>   

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [AllowNull()]
        [Object[]]$InputObject,

        [Parameter(Mandatory=$true)]    
        [hashtable]$Exclude,

        [Parameter(Mandatory=$true)]
        [hashtable]$Where
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $Results = @()


    # FUNCTION MAIN CODE

    # If we have 'Exclude' expressions we process them first...
    if ($Exclude.Count -gt 0)
    {
        foreach ($Object in $InputObject)
        {
            $HasObjectToBeExcluded = $false
            foreach ($Key in $Exclude.Keys)
            {
                $Attribute = $Key.ToString()
                if ($Object."$Attribute" -match $Exclude[$Key])
                {
                    $HasObjectToBeExcluded = $true
                    break
                }
            }

            if ($HasObjectToBeExcluded) { continue }
            $Results  += $Object

        }# End of foreach all objects.

    }# ...otherwise we pass through the input objects.
    else { $Results += $InputObject }
        
    # Without a Where-Object statement we could finish here.
    if (-not $Where.Arguments) 
    {
        Write-Output $Results
        return    
    }

    # If we have a Where-Object command we invoke it here.
    if ($Where.HasScriptblock) { $InvokeCommand = '$Results | Where-Object -FilterScript ' + "$($Query.Where)"}
    else                       { $InvokeCommand = '$Results | Where-Object ' + "$($Query.Where)" }
    $Results = Invoke-Expression -Command $InvokeCommand -ErrorAction Stop 

    Write-Output $Results

}# End of function ExcludeObject

function GetDomainAndContainerName
{
<#
.DESCRIPTION
    The function collects several domain and OU names to support the check and execution
    of the dynamic group queries. To allow the user to keep the dynamic group definition in the info attribute
    of a group as short as possible it is possible to use shortened spellings, but they must be translated into
    correct distinguished names e.g. /Administration stands for the canonicalName hannover-re.grp/Administration
    and this will be replaced with the distinguishedName value.
    The function also get the SID and check for the 'Recycle Bin' feature.
    
.INPUTS
    None
    
    You cannot pipe input to this function.

.OUTPUTS
    System.Collections.Hashtable
    The domain and container data. 

#>

    # PARAMETERS
    [CmdletBinding()]
    param ()

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        [Object[]]$Containers = @()
        $Result               = @{
                                'DistinguishedName'       = $null
                                'Name'                    = $null
                                'DNSName'                 = $null
                                'DNofOUName'              = @{}       # Hashtable with Key/Value (OU Name attribute/OU DN attribute).
                                'DNofCanonicalName'       = @{}       # Hashtable with Key/Value (OU canonicalName attribute/OU DN attribute).
                                'AmbiguousOUNames'        = @()       # All OU 'Name' attribute values which exist multiple times in the domain.
                                'OUDistinguishedNames'    = @()       # All DNs of the OUs.
                                'OUCanonicalNames'        = @()       # All canonicalNames e.g. 'test.com/test1 OU/test2 OU'.
                                'OUReverseCanonicalNames' = @()       # All reverse canonicalNames e.g. 'test2 OU/test1 OU/test.com'.
                                'SID'                     = $null
                                'IsRecycleBinEnabled'     = $false    # Flag if the AD optional feature 'Recycle Bin' has been enabled.
                                 }
        $ExcludedContainers   = @(                                    # We do not allow to use this containers or sub-containers of them.
                                'CN=Keys',
                                'CN=ForeignSecurityPrincipals',
                                'CN=LostAndFound',
                                'CN=Managed Service Accounts',
                                'CN=OperationsManager',
                                'CN=OpsMgrLatencyMonitors',
                                'CN=Program Data',
                                'CN=System',
                                'CN=NTDS Quotas',
                                'CN=TPM Devices',
                                'CN=Microsoft Exchange System Objects'
                                 )

    # FUNCTION MAIN CODE

    $Domain                   = Get-ADDomain
    $Result.Name              = $Domain.Name
    $Result.DNSName           = $Domain.DNSRoot
    $Result.SID               = $Domain.DomainSID.Value
    $Result.DistinguishedName = (Get-ADRootDSE).rootDomainNamingContext

    # Recycle Bin is a special case.
    $Result.DNofOUName.'Recycle Bin' = 'Recycle Bin'

    # First we get all containers and filter out some system containers which we want to exclude for security reasons.
    $Objects = Get-ADObject -Filter "objectClass -eq 'container'" -Properties canonicalName
    foreach ($Object in $Objects)
    {
        $HasObjectExcluded = $false
        foreach ($ExcludedContainer in $ExcludedContainers)
        {
            if ($Object.distinguishedName -match "$ExcludedContainer,$($Result.DistinguishedName)$")
            {
                $HasObjectExcluded = $true
                break    
            }
        }
        if ($HasObjectExcluded) { continue }
        $Containers += $Object
    }

    $Containers += Get-ADOrganizationalUnit -Filter * -Properties canonicalName
    foreach($Container in $Containers)
    {
        # We produce a reverse canonicalName (to protect the spaces we must replace it through '=' before we do the split into an array).
        $CanonicalNameParts   = ($Container.canonicalName -replace ' ','=') -split '/'
        $ReverseCanonicalName = ("$($CanonicalNameParts[-1..-($CanonicalNameParts.Length)])" -replace ' ','/') -replace '=',' '

        $Result.OUDistinguishedNames    += $Container.DistinguishedName
        $Result.OUCanonicalNames        += $Container.canonicalName
        $Result.OUReverseCanonicalNames += $ReverseCanonicalName

        $Result.DNofCanonicalName.Add($Container.canonicalName,$Container.DistinguishedName)
        $Result.DNofCanonicalName.Add($ReverseCanonicalName   ,$Container.DistinguishedName)

        if (-not $Result.DNofOUName.ContainsKey($Container.Name))       { $Result.DNofOUName.Add($Container.Name,$Container.DistinguishedName) }      
        elseif ($Result.AmbiguousOUNames -notcontains $Container.Name ) { $Result.AmbiguousOUNames += $Container.Name }
    }

    $ADFeature = Get-ADOptionalFeature -Filter "Name -eq 'Recycle Bin Feature'"
    if ($ADFeature.EnabledScopes) { $Result.IsRecycleBinEnabled = $true }

    Write-Output $Result

}# End of function GetDomainAndContainerName

function GetDynamicGroup
{
<#
.DESCRIPTION
    The function either get the group defined in the 'Identity' parameter or
    try to find the groups defined through the data in the config file.
    
.INPUTS
    System.String
    An identity value of a domain group.
    
    You cannot pipe input to this function.

.OUTPUTS
    Microsoft.ActiveDirectory.Management.ADGroup
    Array with Active Directory group objects which have been found.

#>

    # PARAMETERS
    [CmdletBinding()]
    param()

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $Results = @()      # All detected groups.
    

    # FUNCTION MAIN CODE

    # Get the groups either from the 'Identity' parameter or from the config file.
    if ($Identity) { $Results += Get-ADGroup -Identity $Identity -Properties info,description -ErrorAction Stop }
    else
    { 
        $SharedParameters = @{Filter = "Name -like '$($Config.NamePattern)'"; Properties = @('info','description')}
        if ($Config.Container) { $Results += Get-ADGroup @SharedParameters -SearchBase $Config.Container -SearchScope Subtree }
        else                   { $Results += Get-ADGroup @SharedParameters }
        
        # Add the statically defined groups from the config file.
        foreach ($GroupIdentity in $Config.Groups)
        {
            try   { $Results += Get-ADGroup -Identity $GroupIdentity -Properties info,description -ErrorAction Stop }
            catch { LogResult -Message "Could not find/read group ($ConfigIdentity)." -Type ERROR }
        }
    }

    # Add the RID as a property to each group object.
    foreach ($Result in $Results)
    {
        $RID = $Result.SID.Value.Substring($Result.SID.Value.LastIndexOf('-') + 1)
        $Result | Add-Member -NotePropertyName RIDN -NotePropertyValue $RID -Force
    }

    Write-Output $Results

}# End of function GetDynamicGroup

function GetDynamicGroupMember
{
<#
.DESCRIPTION
    Get all user or computer samAccountNames which are matching the defined query.
    
.INPUTS
    System.Collections.Hashtable
    The query definition.
    
    You cannot pipe input to this function.

.OUTPUTS
    System.Collections.Hashtable
    The dynamic members and the status code.

.PARAMETER Query
    The query parameter to find the dynamic group members.       
#>

    # PARAMETERS
    [CmdletBinding()]
    param ([hashtable]$Query)

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $Result   = @{
                    'IsValid'         = $true       # Flag if the query has produced failures.
                    'IsEmpty'         = $false      # Flag to show if we have any results.
                    'SamAccountNames' = @()         # All samAccountNames found with this query.
                    'ObjectOf'        = @{}         # Map a samAccountName to its object.
                    }
    
    # FUNCTION MAIN CODE

    <# Fist we check the existence of the groups defined in the 'VMember' parameters.
    While checking we create a hashtable, the Disjunctive Normal Form ($DNF). This structure is used to be filled
    with the group samAccountName (keys) and all group member samAccountNames (values) and doing the 
    union and intersections of the members.
    The following example shows that if an samAccountName is member of Group1 OR (Group2 AND Group3)
    it will be included to the dynamic group.
    $DNF = @{
            'Group1'        = @(Members)
            'Group2,Group3' = @{
                                'Group2'        = @(Members)
                                'Group3'        = @(Members)
                                'SmallestGroup' = Group2
                                'Intersection'  = @(Members which are in Group2 AND Group3)
                                }
            }
    #>
    $DNF     = @{}
    $DNFKeys = @()
    foreach ($VMember in $Query.VMembers)
    {
        # We transform a VMember parameter value in a version without the RID numbers and with the current correct samAccountName of the groups.
        $VMemberWithoutRID = $null  

        # In one VMember parameter line we could have more than one group because then all this groups define an intersection.
        # e.g. VMember:Group1,Group2,Group3 or VMember:Group1[1200],Group2,Group3
        $VMemberParts = $VMember -split ','
        foreach ($VMemberPart in $VMemberParts)
        {
            # VMember group definitions could include a group RID e.g. 'Groupname[RID]'.
            $VMemberPartWithoutRID = $VMemberPart -replace '\[([0-9]*)\]',''      # This is the group name without RID
            try
            {
                # If we have a RID we search the group with the RID.
                if ($VMemberPart -match '(.*?)\[([0-9]+)\]') { $Identity = "$($Domain.SID)-$($Matches[2])" }
                else                                         { $Identity = $VMemberPart }

                $FoundGroup         = Get-ADGroup -Identity $Identity -ErrorAction Stop
                $VMemberWithoutRID += $FoundGroup.samAccountName + ','
                
                # The 'Domain Users' group is not allowed to be a VMember.
                if ($FoundGroup.SID.Value -eq "$($Domain.SID)-513")
                {
                    $script:LogResults += "${Now}:ERROR:$($Group.Name):'Domain Users' group is not allowed to be a VMember."
                    $Result.IsValid = $false
                }

                # If the name of the group which we have found is different from the VMember parameter we log that as information.
                if ($FoundGroup.samAccountName -ne $VMemberPartWithoutRID)
                { $script:LogResults += "${Now}:Info:$($Group.Name):VMember group ($VMemberPart) has been renamed, please change the Query Group definition." }
            }
            catch
            {
                $script:LogResults += "${Now}:ERROR:$($Group.Name):VMember group ($VMemberPart) could not be found."
                $Result.IsValid = $false
            }
        }# End of foreach all groups names in the VMember parameter.

        if (-not $Result.IsValid) { continue }
        
        # Remove the trailing comma and define the DNFKeys which are equal to the 'VMember' parameters
        # but without the [RID] part and with the current group name.
        $VMemberWithoutRID = $VMemberWithoutRID.Trim(',')
        $DNFKeys          += $VMemberWithoutRID

        # Add the VMember group names to the DNF.
        if ($VMemberParts.Count -eq 1) { $DNF.Add($VMemberWithoutRID,$null) }
        else
        {
            $DNF.Add($VMemberWithoutRID,@{SmallestGroup = $null})
            $VMemmberWithoutRIDParts = $VMemberWithoutRID -split ','
            foreach ($VMemmberWithoutRIDPart in $VMemmberWithoutRIDParts) { ($DNF."$VMemberWithoutRID").Add($VMemmberWithoutRIDPart,$null) }
        }

    }# End of foreach all VMember parameters.

    # So if we could not find all 'VMember' groups we cancel the query.
    if ($Result.IsValid -eq $false)
    {
        Write-Output $Result
        return 
    }

    # Here we search in all defined SearchBases for the objects.
    # Because we have four possible types of queries (User,Computer) X (Filter,LDAPFilter) we construct a string 'TypeFilterPair'
    # to avoid too many code nestings.
    if ($Query.Filter) { $TypeFilterPair = "$($Query.Type),Filter" }
    else               { $TypeFilterPair = "$($Query.Type),LDAPFilter"}

    foreach ($SearchBase in $Query.SearchBases) 
    {
        try
        {
            # We have four different ways to get the members.
            switch ($TypeFilterPair)
            {
                'User,Filter'
                { [object[]]$Objects = Get-ADUser -Filter $Query.Filter -Properties $Query.Attributes -SearchBase $SearchBase -SearchScope $Query.SearchScope -ErrorAction Stop }
        
                'User,LDAPFilter'
                { [object[]]$Objects = Get-ADUser -LDAPFilter $Query.LDAPFilter -Properties $Query.Attributes -SearchBase $SearchBase -SearchScope $Query.SearchScope -ErrorAction Stop }
        
                'Computer,Filter'
                { [object[]]$Objects = Get-ADComputer -Filter $Query.Filter -Properties $Query.Attributes -SearchBase $SearchBase -SearchScope $Query.SearchScope -ErrorAction Stop }
        
                'Computer,LDAPFilter'
                { [object[]]$Objects = Get-ADComputer -LDAPFilter $Query.LDAPFilter -Properties $Query.Attributes -SearchBase $SearchBase -SearchScope $Query.SearchScope -ErrorAction Stop }
            }
        }
        catch
        {
            $script:LogResults += "${Now}:ERROR:$($Group.Name):Query in OU ($SearchBase) returns with exception ($($_.Exception.Message))."
            $Result.IsValid     = $false
            Write-Output $Result
            return
        }

        # Here we select the objects from the SearchBase search which do not match with any of the exclude patterns.
        [object[]]$SearchBaseObjects = @()
        $WhereCommand = @{Arguments = $Query.Where; HasScriptblock = $Query.HasScriptblock}
        try { $SearchBaseObjects = ExcludeObject -InputObject $Objects -Exclude $Query.Excludes -Where $WhereCommand -ErrorAction Stop }
        catch
        {
            $script:LogResults += "${Now}:ERROR:$($Group.Name):Exclude operation failed ($($_.Exception.Message))."
            $Result.IsValid     = $false
            Write-Output $Result
            return
        }

        # We update our object index...
        foreach ($SearchBaseObject in $SearchBaseObjects) { $Result.ObjectOf."$($SearchBaseObject.samAccountName)" = $SearchBaseObject }

        # ...and our result hashtable.
        $Result.SamAccountNames += $SearchBaseObjects | Select-Object -ExpandProperty samAccountName

    }# End of foreach all SearchBases.


    # Here we get the objects from the groups defined in the 'VMember' parameters and save it in the DNF structure.
    foreach ($VMember in $DNFKeys)
    {
        $VMemberParts = $VMember -split ','
        
        # If we have only one part we have a simple OR condition... 
        if($VMemberParts.Count -eq 1)
        {
            $Objects = @()
            try
            {
                # First we get the group members...
                [object[]]$GroupObjects = Get-ADGroupMember -Identity $VMemberParts[0] -ErrorAction Stop | Where-Object -FilterScript {$_.objectClass -eq $Query.Type}
                
                # ...then we must get the object once again to ensure the object contains the attributes defined in the 'Exclude' parameters. 
                foreach ($GroupObject in $GroupObjects)
                {
                    switch ($Query.Type)
                    {
                        'User'     { $Object = Get-ADUser     -Identity $GroupObject.samAccountName -Properties $Query.Attributes -ErrorAction Stop }
                        'Computer' { $Object = Get-ADComputer -Identity $GroupObject.samAccountName -Properties $Query.Attributes -ErrorAction Stop }
                    }
                    $Objects += $Object
                }
            }
            catch
            {
                $script:LogResults += "${Now}:ERROR:$($Group.Name):VMember ($VMember) query returns with exception ($($_.Exception.Message))."
                $Result.IsValid     = $false
                Write-Output $Result
                return
            }

            # Here we select the objects from the 'VMember' search which do not match with any of the exclude patterns.
            [object[]]$VMemberObjects = @()
            $WhereCommand = @{Arguments = $Query.Where; HasScriptblock = $Query.HasScriptblock}
            try { $VMemberObjects = ExcludeObject -InputObject $Objects -Exclude $Query.Excludes -Where $WhereCommand -ErrorAction Stop }
            catch
            {
                $script:LogResults += "${Now}:ERROR:$($Group.Name):Exclude operation failed ($($_.Exception.Message))."
                $Result.IsValid     = $false
                Write-Output $Result
                return
            }
            
            # We update our object index ...
            foreach ($VMemberObject in $VMemberObjects)
            {
                # Skip the duplicates.
                if ($Result.ObjectOf.ContainsKey($VMemberObject.samAccountName)) { continue }

                $Result.ObjectOf."$($VMemberObject.samAccountName)" = $VMemberObject
            }

            # ...and samAccountNames in the DNF.
            $DNF."$VMember" = @($VMemberObjects | Select-Object -ExpandProperty samAccountName)
        }
        # ...otherwise we have an AND condition.
        else
        {
            # Here we do the intersection operation for all groups in the VMember parameter...
            $SmallestGroupCount = 1000000000000000000
            $DNF."$VMember".'Intersection' = @()
            $DNF."$VMember".'ObjectOf'     = @{}
            
            foreach ($VMemberPart in $VMemberParts)
            {
                $Objects = @()
                try
                {
                    # First we get the group members...
                    [object[]]$GroupObjects = Get-ADGroupMember -Identity $VMemberPart -ErrorAction Stop | Where-Object -FilterScript {$_.objectClass -eq $Query.Type}
                    
                    # ...then we must get the object once again to ensure the object contains the attributes defined in the 'Exclude' parameters. 
                    foreach ($GroupObject in $GroupObjects)
                    {
                        switch ($Query.Type)
                        {
                            'User'     { $Object = Get-ADUser     -Identity $GroupObject.samAccountName -Properties $Query.Attributes -ErrorAction Stop }
                            'Computer' { $Object = Get-ADComputer -Identity $GroupObject.samAccountName -Properties $Query.Attributes -ErrorAction Stop }
                        }
                        $Objects += $Object
                    }
                }
                catch
                {
                    $script:LogResults += "${Now}:ERROR:$($Group.Name):VMember ($VMember) query returns with exception ($($_.Exception.Message))."
                    $Result.IsValid     = $false
                    Write-Output $Result
                    return
                }
                
                # Here we select the objects from the SearchBase search which do not match with any of the exclude patterns.
                [object[]]$VMemberObjects = @()
                $WhereCommand = @{Arguments = $Query.Where; HasScriptblock = $Query.HasScriptblock}
                try { $VMemberObjects = ExcludeObject -InputObject $Objects -Exclude $Query.Excludes -Where $WhereCommand -ErrorAction Stop }
                catch
                {
                    $script:LogResults += "${Now}:ERROR:$($Group.Name):Exclude operation failed ($($_.Exception.Message))."
                    $Result.IsValid     = $false
                    Write-Output $Result
                    return
                }
                
                # We update the AND object index...
                foreach ($VMemberObject in $VMemberObjects)
                { $DNF."$VMember".ObjectOf."$($VMemberObject.samAccountName)" = $VMemberObject }

                # ...and samAccountNames in the DNF.
                $DNF."$VMember"."$VMemberPart" = @($VMemberObjects | Select-Object -ExpandProperty samAccountName)
                
                # Check if this group count is the smallest.
                if ($VMemberObjects.Count -lt $SmallestGroupCount)
                {
                    $SmallestGroupCount           = $VMemberObjects.Count
                    $DNF."$VMember".SmallestGroup = $VMemberPart
                }

            }# End of foreach all parts of a VMember parameter.

            # ...and then we check which members in the smallest group are also in all other groups (intersection).
            $SmallestGroupName = $DNF."$VMember".SmallestGroup
            foreach ($SamAccountName in $DNF."$VMember"."$SmallestGroupName")
            {
                $IsSamAccountNameInAllGroups = $true
                foreach ($VMemberPart in $VMemberParts)
                {
                    if ($VMemberPart -eq $SmallestGroup) { continue }
                    if ( $DNF."$VMember"."$VMemberPart" -notcontains $SamAccountName)
                    {
                        $IsSamAccountNameInAllGroups = $false
                        break
                    }
                }
                
                # Here we add the elements to the intersection and to the 'Result' object.
                if ($IsSamAccountNameInAllGroups)
                {
                    $DNF."$VMember".Intersection       += $SamAccountName

                    if ($Result.ObjectOf.ContainsKey($SamAccountName)) { continue }
                    $Result.ObjectOf["$SamAccountName"] = $DNF."$VMember".ObjectOf["$SamAccountName"]
                } 

            }# End of foreach all samAccountNames in the smallest Group.

        }# End of the AND condition.
    }# End of foreach all VMember keys in the DNF.

    # Now we add the samAccountNames which we have found into the 'Result' object.
    foreach ($VMember in $DNFKeys)
    {
        if ($DNF."$VMember" -is [array]) { $Result.SamAccountNames += $DNF."$VMember" }
        else                             { $Result.SamAccountNames += $DNF."$VMember".Intersection }
    }

    # We should remove the elements that are duplicate...
    $Result.SamAccountNames = @($Result.SamAccountNames | Select-Object -Unique)
    
    # ...and the samAccountNames must be run through the transform process. We need to check if this new names exist
    # in the Active Directory.
    if ($Query.SamIn)
    {
        [string[]]$TransformedNames = TransformSamAccountName -Name $Result.SamAccountNames -Pattern $Query.SamIn -Transformation $Query.SamOut
        $Result.samAccountNames = @()
        $Result.ObjectOf        = @{}
        foreach ($TransformedName in $TransformedNames)
        {
            if ($Result.samAccountNames -contains $TransformedName) { continue }
            # Check if an object with the transformed samAccountName exist.
            try
            {
                $Object = Get-ADObject -Filter "samAccountName -eq '$TransformedName'" -Properties $Query.Attributes -ErrorAction Stop
                $Result.samAccountNames += $TransformedName
                $Result.ObjectOf.Add($TransformedName,$Object)
            }
            catch
            {
                LogResult -Message "The transformed samAccountName ($TransformedName) could not be found." -GroupLog -Type ERROR
                continue
            }
        }# End of foreach all transformed names.
    }# End of samAccountName transformation.

    if($Result.samAccountNames.Count -eq 0) { $Result.IsEmpty = $true }

    Write-Output $Result

}# End of function GetDynamicGroupMember

function LogResult
{
<#
.DESCRIPTION
    The function logs and outputs an error message to the host.    
    
.INPUTS
    System.String
    The message to log.
    
    You cannot pipe input to this function.

.OUTPUTS
    None

.PARAMETER Message
    The text message which should be logged into the log file and output to the host.

.PARAMETER GroupLog
    Add the message to the GroupLog instead of the ScriptLog.

.PARAMETER Type
    The type of the message to be logged.
    The acceptable values for this parameter are: ERROR, INFO or WHATIF

#>

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string]$Message,

        [Parameter(Mandatory=$true)]
        [ValidateSet('ERROR','INFO','WHATIF')]
        [string]$Type,

        [switch]$GroupLog
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # FUNCTION ARGUMENTS
        $TypeArgument    = $Type.ToUpper()
        $MessageArgument = $Message

        # VARIABLES
        # NA


    # FUNCTION MAIN CODE

    if ($GroupLog)
    {
        $script:GroupLogResults += "${Now}:${HostInstance}:${TypeArgument}:$Message"
        $MessageArgument         = "`t$Message"
    }
    else { $script:LogResults += "${Now}:${TypeArgument}:$Message" }

    switch ($Type)
    {
        'ERROR'  { $MessageColor = 'Red' }
        'INFO'   { $MessageColor = 'Yellow' }
        'WHATIF' { $MessageColor = 'White' }
    }
    
    # Collect all groups with errors.
    if ($Group -and ($Type -eq 'ERROR')) { $script:FailedGroups += $Group }
       
    Write-Host $MessageArgument -ForegroundColor $MessageColor

}# End of function LogResult

function MoveObjectToContainer
{
<#
.DESCRIPTION
    This function moves all newly added members of a dynamic group into the
    configured PushOU or if the members have been removed from the group into the
    PullOU.
    
.INPUTS
    System.String
    User or Computer samAccountName.

    System.String
    Destination and source path.

    You cannot pipe input to this function.

.OUTPUTS
    None

.NOTES
    The function supports the pseudo-container 'Recycle Bin', in this case a member object
    will be removed from the domain.

.PARAMETER Identity
    List of user or computer objects which should be moved to another container.
    The acceptable value for this parameter is the samAccountName.
 
.PARAMETER Destination
    The distinguishedName of the destination container.

.PARAMETER Source
    The distinguishedName of the PushOU.

.PARAMETER Scope
    Defines the size of the PushOU.
    The acceptable values for this parameter are: OneLevel or Subtree.

#>   

    # PARAMETERS
    [CmdletBinding(SupportsShouldProcess)]
    param
    (
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [string[]]$Identity,

        [Parameter(Mandatory=$true)]    
        [string]$Destination,

        [string]$Source,

        [Parameter(Mandatory=$true)]
        [ValidateSet('OneLevel','Subtree')]
        [string]$Scope
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        # NA


    # FUNCTION MAIN CODE

    foreach ($SamAccountName in $Identity)
    {
        $MemberObject = $QueryResults.ObjectOf[$SamAccountName]
        if (-not $MemberObject) { $MemberObject = Get-ADObject -Filter "samAccountName -eq '$SamAccountName'" -Properties 'msDS-parentdistname'  }
        $MemberContainer = $MemberObject.'msDS-parentdistname'
        $MemberType      = $MemberObject.'ObjectClass'
        
        # Check if object is already in 'Destination' container or not anymore in the 'Source' container.
        $HasObjectToBeMoved = $false
        if ($Source)
        {
            # We move the object only if it is still in the 'Source' container...
            switch ($Scope)
            {
                'OneLevel'
                { if ($MemberContainer -ne $Source) { $HasObjectToBeMoved= $true } }
                'Subtree'
                { if (-not $MemberContainer.EndsWith($Source,'CurrentCultureIgnoreCase')) { $HasObjectToBeMoved= $true } }
            }
            
            # ...and the object is not already in the 'Destination' container.
            if ($MemberContainer -ne $Destination) { $HasObjectToBeMoved = $true }
        
        }
        elseif ($Destination -eq 'Recycle Bin') { $HasObjectToBeMoved = $true }
        else 
        {
            # We move the object only if it is not already in the 'Destination' container.
            switch ($Scope)
            {
                'OneLevel'
                { if ($MemberContainer -ne $Destination) { $HasObjectToBeMoved= $true } }
                'Subtree'
                { if (-not $MemberContainer.EndsWith($Destination,'CurrentCultureIgnoreCase')) { $HasObjectToBeMoved= $true } }
            }
        }# End of test whether the object has to be moved.

        # Skip the object if necessary.
        if (-not $HasObjectToBeMoved) { continue }

        # Write the first character of 'user|computer' type in capital letter.
        $null = $MemberType -match '^(u|c)'
        $Type = $MemberType -replace '^(u|c)',$Matches[1].ToUpper()

        # Do the move or deletion of the object.
        if ($PSCmdlet.ShouldProcess($Destination,"Update objects in OU."))
        {
            try
            {
                $SharedParameters = @{Identity=$MemberObject; WhatIf=$false; ErrorAction='Stop'}
                if ($Destination -eq 'Recycle Bin') { Remove-ADObject @SharedParameters }
                else                                { Move-ADObject   @SharedParameters -TargetPath $Destination }
                $EventLogType = "INFO"
            }
            catch
            {
                LogResult -Message "$($Group.Name):Cannot move object ($SamAccountName) to destination ($Destination) ($($_.Exception.Message))." -Type ERROR
                $EventLogType = "ERROR"
                throw $_
                break
            }    
        }
        else { $EventLogType = "WHATIF" }
        
        LogResult -Message "Move $Type ($SamAccountName) from ParentOU ($MemberContainer)) to destination ($Destination)." -Type $EventLogType -GroupLog

    }# End of foreach all members to be moved.

}# End of function MoveObjectToContainer

function SendMail
{
<#
.DESCRIPTION
    Send the group changes which have been done to the subscriber.        
    
.INPUTS
    System.String
    List of receiver addresses.

    You cannot pipe input to this function.

.OUTPUTS
    None

.PARAMETER Receiver
    List of receiver mail addresses.

#>   

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [string[]]$Receiver,

        [Parameter(Mandatory=$true)]
        [string]$Subject
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $MailBody = @()


    # FUNCTION MAIN CODE

    # Create the Body string and send the mail.
    $Parameters = @{From       = $Config.Mail.From
                    To         = $Receiver
                    Subject    = $Subject
                    Body       = $null
                    SmtpServer = $Config.Mail.SmtpServer
                    Port       = 25
                   }
    
    $MailBody += "Date: $Today"
    $MailBody += "Host: $($Env:COMPUTERNAME)"
    $MailBody += "Instance: $($Config.InstanceID)"
    $MailBody += "Group: $($Group.SamAccountName)"
    $MailBody += "RID: $($Group.RIDN)"
    $MailBody += "Description: $($Group.Description)"
    $MailBody += "`r`n`r`n"
    $MailBody += "--------------------------------------------------------------------------------"
    $MailBody += "The following updates were made to the group:"
    $MailBody += "`r`n"
    foreach ($Line in $GroupLogResults) { $MailBody += $Line }
    $MailBody += "`r`n"
    $MailBody += "--------------------------------------------------------------------------------"

    $Parameters.Body = $MailBody
    if ($Config.Mail.Port) { $Parameters.Port = $Config.Mail.Port }

    Send-MailMessage @Parameters
        
}# End of function SendMail.

function TestConfigurationData
{
<#
.DESCRIPTION
    The function checks all data types and values of the configuration json file.    
    
.INPUTS
    System.Management.Automation.PSCustomObject
    The object containing all configuration data.
    
    You cannot pipe input to this function.

.OUTPUTS
    None

.PARAMETER InputObject
    A PSCustomObject with all configuration data of the script.
#>

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [PSCustomObject]$InputObject
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $ConfigParameters  = @{}       # Hashtable with all parameters found in the Config file.
        $HasErrorsFound    = $false
        $HasMailErrorFound = $false
        $MailTestResults   = @()       # All errors in the 'Mail' section
        $TestResults       = @()       # All errors found by the test.
        

    # FUNCTION MAIN CODE

    # Create a hashtable with all parameters found in the Config file.
    # Key is the parameter name and the value is the PSNoteProperty object.
    foreach ($Property in $InputObject.psobject.Properties)
    { if ($null -ne $Property.Value) { $ConfigParameters.Add($Property.Name,$Property) } }

    # Test all parameters of type System.Boolean.
    $BooleanParameters = @('WhatIf','IsRootSearchAllowed','IsScriptblockAllowed','ForceRecycleBin','EnableMail')
    foreach ($BooleanParameter in $BooleanParameters)
    {
        if (-not $ConfigParameters.ContainsKey($BooleanParameter))
        {
            $TestResults   += "${Now}:ERROR:Missing parameter ($BooleanParameter)."
            $HasErrorsFound = $true
        }
        elseif ($ConfigParameters[$BooleanParameter].TypeNameOfValue -ne 'System.Boolean')
        {
            $TestResults   += "${Now}:ERROR:Parameter ($BooleanParameter) must be of type (Boolean)."
            $HasErrorsFound = $true
        }
    }

    # Test parameter 'InstanceID'.
    if ($ConfigParameters.ContainsKey('InstanceID'))
    {
        if($ConfigParameters['InstanceID'].TypeNameOfValue -ne 'System.String')
        {
            $TestResults   += "${Now}:Parameter (InstanceID) must be of type (String)."
            $HasErrorsFound = $true
        }
        elseif($InputObject.InstanceID -notmatch $VALID_INSTANCEID_REGEX)
        {
            $TestResults   += "${Now}:Parameter (InstanceID) has not a valid format."
            $HasErrorsFound = $true
        }
    }

    # Test parameter 'NamePattern'.
    if (-not $ConfigParameters.ContainsKey('NamePattern'))
    {
        $TestResults   += "${Now}:ERROR:Missing parameter (NamePattern)."
        $HasErrorsFound = $true
    }
    elseif ($ConfigParameters['NamePattern'].TypeNameOfValue -ne 'System.String')
    {
        $TestResults   += "${Now}:ERROR:Parameter (NamePattern) must be of type (String)."
        $HasErrorsFound = $true
    }

    # Test parameter 'LogPath'
    if (-not $ConfigParameters.ContainsKey('LogPath'))
    {
        $TestResults   += "${Now}:ERROR:Missing parameter (LogPath)."
        $HasErrorsFound = $true
    }
    elseif ($ConfigParameters['LogPath'].TypeNameOfValue -ne 'System.String')
    {
        $TestResults   += "${Now}:ERROR:Parameter (LogPath) must be of type (String)."
        $HasErrorsFound = $true
    }
    elseif (-not (Test-Path -Path $InputObject.LogPath -PathType Container))
    {
        $TestResults   += "${Now}:ERROR:Path of parameter (LogPath) could not be found."
        $HasErrorsFound = $true
    }

    # Test parameter 'Container'.
    if ($ConfigParameters.ContainsKey('Container'))
    {
        if($ConfigParameters['Container'].TypeNameOfValue -ne 'System.String')
        {
            $TestResults   += "${Now}:ERROR:Parameter (Container) must be of type (String)."
            $HasErrorsFound = $true
        }
        elseif ($InputObject.Container.Length -gt 0)
        {
            try { $null = Get-ADOrganizationalUnit -Identity $InputObject.Container -ErrorAction Stop }
            catch
            {
                $TestResults   += "${Now}:ERROR:Path of parameter (Container) could not be found."
                $HasErrorsFound = $true
            }
        }
    }

    # Test parameter 'Groups'.
    if ($ConfigParameters.ContainsKey('Groups'))
    {
        if($ConfigParameters['Groups'].TypeNameOfValue -ne 'System.Object[]')
        {
            $TestResults   += "${Now}:ERROR:Parameter (Groups) must be of type (Array)."
            $HasErrorsFound = $true
        }
    }

    # Test parameter 'LDAPAliases'.
    if (-not $ConfigParameters.ContainsKey('LDAPAliases'))
    {
        $TestResults   += "${Now}:ERROR:Missing parameter (LDAPAliases)."
        $HasErrorsFound = $true
    }
    elseif ($ConfigParameters['LDAPAliases'].TypeNameOfValue -ne 'System.Management.Automation.PSCustomObject')
    {
        $TestResults   += "${Now}:ERROR:Parameter (LDAPAliases) must be of type (Hashtable)."
        $HasErrorsFound = $true
    }

    # Test parameter 'Mail'
    if (-not $ConfigParameters.ContainsKey('Mail'))
    {
        $MailTestResults  += "${Now}:ERROR:Missing parameter (Mail)."
        $HasMailErrorFound = $true
    }
    elseif ($ConfigParameters['Mail'].TypeNameOfValue -ne 'System.Management.Automation.PSCustomObject')
    {
        $MailTestResults  += "${Now}:ERROR:Parameter (Mail) must be of type (Hashtable)."
        $HasMailErrorFound = $true    
    }
    else
    {
        # Test parameter 'Mail.SmtpServer'
        if ($InputObject.Mail.SmtpServer)
        {
            if ($InputObject.Mail.SmtpServer -isnot [string])
            {
                $MailTestResults  += "${Now}:ERROR:Parameter (Mail.SmtpServer) must be of type (String)."
                $HasMailErrorFound = $true    
            }
        }
        else 
        {
            $MailTestResults  += "${Now}:ERROR:Missing parameter (Mail.SmtpServer)."
            $HasMailErrorFound = $true
        }

        # Test parameter 'Mail.Port'
        if ($InputObject.Mail.Port -and ($InputObject.Mail.Port -isnot [int]))
        { 
            $MailTestResults  += "${Now}:ERROR:Parameter (Mail.Port) must be of type (Integer)."
            $HasMailErrorFound = $true    
        }       
        
        # Test parameter 'Mail.From'
        if ($InputObject.Mail.From)
        {
            if ($InputObject.Mail.From -isnot [string])
            {
                $MailTestResults  += "${Now}:ERROR:Parameter (Mail.From) must be of type (String)."
                $HasMailErrorFound = $true    
            }
        }
        else 
        {
            $MailTestResults  += "${Now}:ERROR:Missing parameter (Mail.From)."
            $HasMailErrorFound = $true
        }

        if ($null -eq $InputObject.Mail.Domain)
        {
            $MailTestResults  += "${Now}:ERROR:Missing parameter (Mail.Domain)."
            $HasMailErrorFound = $true
        }
        elseif ($InputObject.Mail.Domain.Count -eq 0)
        {
            $MailTestResults  += "${Now}:ERROR:Parameter (Mail.Domain) must include at least one domain name."
            $HasMailErrorFound = $true
        }
    }# End of test parameter 'Mail'.

    if ($InputObject.EnableMail -and $HasMailErrorFound)
    {
        $TestResults   += $MailTestResults
        $HasErrorsFound = $true
    }

    if ($HasErrorsFound)
    {
        $TestResults | Out-File -FilePath "$PSScriptRoot\Error_$Today.txt" -Append -WhatIf:$false
        Write-Error -Message "Processing aborted, the config file ($ConfigFile) contains errors." -Category InvalidData -ErrorAction Stop
    }

}# End of function TestConfigurationData

function TransformSamAccountName
{
<#
.DESCRIPTION
    The function transforms one samAccountName in another.
    
.INPUTS
    System.String
    The samAccountName value to be transformed.
    
    You cannot pipe input to this function.

.OUTPUTS
    System.String
    The transformed samAccountName value.

.PARAMETER Name
    The samAccountName to be transformed.

.PARAMETER Pattern
    An regular expression to capture data which should be used in the transform process.

.PARAMETER Transformation
    The result string which uses the capture groups from the regex pattern.
#>

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory=$true)]
        [AllowEmptyCollection()]
        [string[]]$Name,
    
        [Parameter(Mandatory=$true)]
        [string]$Pattern,

        [Parameter(Mandatory=$true)]
        [string]$Transformation
    )

    # PARAMETER CHECK
    # NA

    # DECLARATIONS AND DEFINITIONS
        # VARIABLES
        $Results = @()          # All transformed samAccountNames.


    # FUNCTION MAIN CODE
    foreach ($SamAccountName in $Name)
    {
        if ($Pattern)
        {
            if ($SamAccountName -notmatch $Pattern) { continue }

            # Run through all matches and replace the capture group identifier [0],[1],... through their values.
            $TransformedName = $Transformation
            foreach ($Index in $Matches.Keys) { $TransformedName = $TransformedName.Replace("[$Index]",$Matches.$Index) }
            $Results += $TransformedName
        }
        else { $Results += $SamAccountName }

    }# End of foreach all submitted names.

    Write-Output $Results

}# End of function TransformSamAccountName


