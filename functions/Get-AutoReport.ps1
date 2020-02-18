##############################
#.SYNOPSIS
#TODO
#
#.DESCRIPTION
#TODO
#
#.PARAMETER AutoRptFile
#TODO
#
#.EXAMPLE
#TODO
#
#.EXAMPLE
#TODO
##############################
function Get-AutoReport {
    [CmdletBinding()]
    param (
        [string]$AutoRptFile
    )

    process {
        # Create new AutoReport object
        $autoReport = [AutoReport]::new($AutoRptFile);
        $autoReport
    }
}