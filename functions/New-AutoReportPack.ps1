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
#.PARAMETER TargetConnectionName
#A reference number which you wish to associate the instantiation of the Analysis Pack.
#
#.EXAMPLE
#TODO
#
#.EXAMPLE
#TODO
##############################
function New-AutoReportPack {
    [CmdletBinding()]
    param (
        [string]$SourceFolder
        ,[string]$AutoRptFile
    )

    process {
        Compress-7Zip -Path $SourceFolder -ArchiveFileName $AutoRptFile;
    }
}