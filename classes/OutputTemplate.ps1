class OutputTemplate {
    [string]$TemplateFileName;
    [byte[]]$TemplateBytes;

    #--------------#
    # Constructors #
    #--------------#
    OutputTemplate([System.IO.FileInfo]$TemplateFile) {
        $this.TemplateFileName = $_.Name;
        $this.TemplateBytes = [System.IO.File]::ReadAllBytes( $TemplateFile.FullName );
    }
}