class Mapping {
    [string]$OutputFileNameTemplate;
    [string]$OutputTemplateName;
    [string]$OutputType;
    [SourceMap[]]$SourceMaps;

    #---------#
    # Methods #
    #---------#
    [string]CalculateOutputFileName([object]$CurrentStateProps){
        
        [string]$cmd = "";

        # Create variables for each Key Value pair
        $CurrentStateProps.Keys | ForEach-Object {
            $cmd += "New-Variable -Name '$($_)' -Value '$($CurrentStateProps.Item($_))';";
        }
        $cmd += "`"$($this.OutputFileNameTemplate)`";";
        
        return Invoke-Expression $cmd;
    }

    [string[]]ExtractOutputFileNameTemplateParameters(){
        return (Select-String -InputObject $this.OutputFileNameTemplate -Pattern "(\$.\w*)" -AllMatches | ForEach-Object{$_.Matches}).Value;
    }
}
