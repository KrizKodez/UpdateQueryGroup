<#PSScriptInfo

.TYPE Private Library

.TEMPLATEVERSION 1

.GUID 0CAF51EE-D0AF-4B0C-95C3-42878C0C5554

.FUNCTIONS
    GetNoLoginSinceDaysExpression
   
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
    2025-07-22,All functions,Christoph Rust,Initial Release

.Description
    This library contains private functions for the scripts defined in REQUIREDSCRIPTS.

#>

# DECLARATIONS AND DEFINITIONS
    # CONSTANTS
    # NA

    # VARIABLES
    $FunctionOfDynamicExpression = @{}

# SCRIPTBLOCKS
# NA


# FUNCTIONS

function GetNoLoginSinceDaysExpression
{
<#
.DESCRIPTION
    The function creates a dynamic expression to be injected in a query parameter.
    It calculates a FILETIME value which is the treshold for the lastLogonTimestamp attribute. 
    
.INPUTS
    System.String
    The number of days.
    
    You cannot pipe input to this function.

.OUTPUTS
    System.String
    The expression is 'lastLogonTimestamp -lt <FILETIME>'.

.PARAMETER Day
    The number of days to find objects where the last logon was before this timespan.       
#>

    # PARAMETERS
    [CmdletBinding()]
    param
    (
        [Parameter(Position=0)]
        [string]$Day
    )

    # DECLARATIONS AND DEFINITIONS
        # FUNCTION ARGUMENTS
        $DayArgument = $null
    
        # VARIABLES
        # NA

    # PARAMETER CHECK
    try   { $DayArgument = [int]$Day }
    catch { throw; return}


    # FUNCTION MAIN CODE

    $Today    = Get-Date
    $Treshold = $Today.AddDays(-$DayArgument).ToFileTime()
    $Result   = "lastLogonTimestamp -lt $Treshold"

    Write-Output $Result

}# End of function GetNoLoginSinceDaysExpression.

# Register the aliases for the dynamic expression.
$FunctionOfDynamicExpression.Add('NLS','GetNoLoginSinceDaysExpression')
$FunctionOfDynamicExpression.Add('NoLoginSinceDays','GetNoLoginSinceDaysExpression')
