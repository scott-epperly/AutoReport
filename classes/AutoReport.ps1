class AutoReport {
    #------------#
    # Properties #
    #------------#
    #[string]$OutputCreationStep;

    [Mapping[]]$Mappings;
    [Connection[]]$Connections;
    [Query[]]$Queries;
    [OutputTemplate[]]$Templates;

    # Local variables
    hidden [System.IO.DirectoryInfo]$tmpDir;

    #--------------#
    # Constructors #
    #--------------#
    AutoReport([string]$AutoRptFile){
        #Create temp folder
        $this.tmpDir = $this.CreateTemporaryDirectory();
        Write-Debug "[AutoReport.Constructor] Created temporary folder: $($this.tmpDir.FullName)";

        #unzip .autorpt file to temp folder
        Expand-7Zip -ArchiveFileName $AutoRptFile -TargetPath $this.tmpDir.FullName;

        #read config.json and assign properties
        Write-Debug "[AutoReport.Constructor] Reading config.json"
        $config = Get-Content (Join-Path $this.tmpDir.FullName "config.json") | ConvertFrom-Json;
        
        #$this.OutputCreationStep = $config.OutputCreationStep;
        
        $this.Mappings = $config.Mappings | ForEach-Object {
            $mapping = $_;
            
            [Mapping]$obj = [Mapping]::new();
            $obj.OutputFileNameTemplate = $mapping.OutputFileNameTemplate;
            $obj.OutputTemplateName = $mapping.OutputTemplateName;
            $obj.OutputType = $mapping.OutputType;

            $obj.SourceMaps = $mapping.SourceMaps | ForEach-Object {
                $sourceMap = $_;

                switch ($mapping.OutputType) {
                    "Excel" {[SourceMapExcel]($sourceMap);}
                    # As more output types come on, add logic here
                }
            } 
            $obj;
        };

        $this.Connections = $config.Connections | ForEach-Object {
            [Connection]$_;
        };

        # Populate Queries
        $this.Queries = Get-ChildItem (Join-Path $this.tmpDir.FullName "queries\*") | ForEach-Object {
            [string]$queryName = $_.Name;
            [bool]$returnQuery = $false;

            [int]$minExecutionOrder = 0;
            [int]$maxExecutionOrder = 0;
            [string]$dataProviderName = "";

            # only load in query files that are referenced in the map
            $this.Mappings | ForEach-Object {
                $i = 1;
                $_.SourceMaps | ForEach-Object {
                    if($_.QueryFileName -eq $queryName) {
                        $returnQuery = $true;
                        $dataProviderName = $_.DataProviderName;
                        
                        if($minExecutionOrder -eq 0) {
                            $minExecutionOrder = $i;
                            $maxExecutionOrder = $i;
                        }
                        else {
                            $maxExecutionOrder = $i;
                        }
                    }
                    $i++;
                }
            }
            
            if($returnQuery){
                [Query]::new($_, $dataProviderName, $minExecutionOrder, $maxExecutionOrder);
            }
        };

        # Populate Templates collection
        $this.Templates = Get-ChildItem (Join-Path $this.tmpDir.FullName "templates\*") | ForEach-Object {
            [OutputTemplate]::new( $_ );
        }

        #Remove the temp folder and all files in it.
        $this.tmpDir.Delete($true);
    }

    #---------#
    # Methods #
    #---------#
    [void]CreateOutput([string]$OutputFileFullName, [string]$TemplateFileName){
        # Get containing folder name and create it if it doesn't exist
        $OutputFolderName = Split-Path $OutputFileFullName -Parent;
        Write-Debug "[AutoReport.CreateOutput()]Checking destination folder: $OutputFolderName";
        if (!(Test-Path $OutputFolderName)) {
            New-Item $OutputFolderName -ItemType Directory;
        }
        
        # Write the template bytes to target file
        try{
            Write-Debug "[AutoReport.CreateOutput()] Creating template base for output: $OutputFileFullName";
            [System.IO.File]::WriteAllBytes($OutputFileFullName, ($this.Templates | Where-Object {$_.TemplateFileName -eq $TemplateFileName}).TemplateBytes);
        }
        catch {
            throw "Error writing file: $($_.Exception.Message)";
        }
    }

    # Excel Output overload
    [void]WriteToOutput([string]$OutputFileFullName, [System.Data.Datatable]$dt, [string]$WorksheetName, [string]$Cell, [string]$TableStyle){
        #Calculate StartColumn & StartRow
        [int]$startColumn = $this.Base26ToDecimal(($Cell -replace "[1-9]"));
        [int]$startRow = $Cell -replace "[A-Z]";
        $tableName = $WorksheetName -replace " ";

        Write-Debug "[AutoReport.WriteToOutput()] startColumn: $startColumn"
        Write-Debug "[AutoReport.WriteToOutput()] startRow: $startRow"
        Write-Debug "[AutoReport.WriteToOutput()] tableName: $tableName"

        if ($dt.Rows.Count -gt 0) {
            Write-Verbose "Writing $($dt.Rows.Count) rows to $WorksheetName starting at cell $Cell"
            $dt | Export-Excel -Path $OutputFileFullName -WorksheetName $WorksheetName -StartRow $startRow -StartColumn $startColumn -TableName $tableName -TableStyle $TableStyle -ExcludeProperty RowError, RowState, Table, ItemArray, HasErrors;
        }
        else {
            Write-Verbose "No rows found to write to $WorksheetName starting at cell $Cell"
        }
    }

    #----------------#
    # Helper Methods #
    #----------------#
    hidden [System.IO.DirectoryInfo]CreateTemporaryDirectory() {
        [string]$parent = [System.IO.Path]::GetTempPath()
        [string]$name = [System.Guid]::NewGuid()
        $ret = New-Item -ItemType Directory -Path (Join-Path $parent "AutoReport\autorpt$name")
        return $ret;
    }

    hidden [int]Base26ToDecimal([string]$Base26) {
        $alphabet = "abcdefghijklmnopqrstuvwxyz"
        $inputarray = $Base26.ToLower().ToCharArray();
        [array]::reverse($inputarray)
        [long]$decNum=0
        $pos=0

        foreach ($c in $inputarray)
        {
            $decNum += ($alphabet.IndexOf($c) + 1) * [long][Math]::Pow(26, $pos)
            $pos++
        }
        return $decNum;
    }
}