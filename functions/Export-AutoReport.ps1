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
function Export-AutoReport {
    [CmdletBinding()]
    param (
        [string]$AutoRptFile
        ,[string[]]$TargetConnectionName
        ,[string]$OutputPath
        ,[object[]]$UserParamHash = @(@{"dummy"="1"})
    )

    process {

        Write-Debug "Parameter - AutoRptFile: $AutoRptFile";
        Write-Debug "Parameter - TargetConnectionName: $TargetConnectionName";
        Write-Debug "Parameter - OutputPath: $OutputPath";
        Write-Debug "Parameter - UserParamHash: $UserParamHash";

        # Initialize Output File Name
        [string]$currentOutputFileFullName;

        # Create new AutoReport object
        $autoReport = [AutoReport]::new($AutoRptFile);

        # Iterate through each TargetConnectionName - current approach is not efficient for connection reuse or parallel processing
        $TargetConnectionName | ForEach-Object {
            $currentTargetConnectionName = $_;
            Write-Verbose "Iterating queries through Connection: $currentTargetConnectionName";
            
            # Get connection object
            [Connection]$currentConnection = $autoReport.Connections | Where-Object{$_.Name -eq $currentTargetConnectionName};
            
            # Iterate through each UserParam Hash
            $UserParamHash | ForEach-Object {
                $currentParamHash = $_;
                Write-Verbose "Iterating queries through user param set";

                #Iterate through all queries defined
                $autoReport.Queries | Sort-Object MinExecutionOrder | ForEach-Object {
                    $currentQuery = $_;

                    Write-Verbose "Executing query file: $($currentQuery.QueryFileName)";
                    
                    # Execute the query under the current Connection & Param context
                    switch ($currentQuery.DataProviderName) {
                        "Kusto" {
                            Write-Debug "Executing Kusto query"
                            [System.Data.Datatable[]]$dt = $currentQuery.ExecuteKQL($currentConnection.ConnectionString, $currentParamHash);
                            break;
                        }
                        default {
                            Write-Debug "Executing SQL query"
                            [System.Data.Datatable[]]$dt = $currentQuery.ExecuteSQL($currentConnection.ConnectionString, $currentQuery.DataProviderName, $currentParamHash);
                            break;
                        }
                    }

                    # Iterate through Mappings
                    $autoReport.Mappings | ForEach-Object {
                        $currentMapping = $_;

                        $currentMapping.SourceMaps | ForEach-Object {
                            $currentSourceMap = $_;
                            
                            if ($currentSourceMap.QueryFileName -eq $currentQuery.QueryFileName)
                            {
                                # Construct current state param hash
                                $currentStateProps = $currentParamHash.Clone();
                                $currentStateProps.Add("ArTargetConnection", $currentConnection.Name);
                                $currentStateProps.Add("ArQuery", [System.IO.Path]::GetFileNameWithoutExtension($currentQuery.QueryFileName));
                                $currentStateProps.Add("ArDataTableId", $currentSourceMap.DataTableId)
                                
                                # Calculate target output file name
                                [string]$OutputFileFullName = Join-Path $OutputPath $currentMapping.CalculateOutputFileName($currentStateProps);

                                # If we haven't yet encountered this name, it's time to create the output file
                                if( $currentOutputFileFullName -ne $OutputFileFullName){
                                    
                                    $currentOutputFileFullName = $OutputFileFullName;

                                    Write-Verbose "Creating output file: $currentOutputFileFullName";
                                    $autoReport.CreateOutput($currentOutputFileFullName, $currentMapping.OutputTemplateName);
                                }

                                # Write DataTables to output
                                #Write-Verbose "Applying $($dt[$currentSourceMap.DataTableId].Rows.Count) rows to output."
                                switch ($currentMapping.OutputType) {
                                    "Excel" {
                                        $autoReport.WriteToOutput([string]$currentOutputFileFullName, $dt[$currentSourceMap.DataTableId], $currentSourceMap.WorksheetName, $currentSourceMap.Cell, $currentSourceMap.TableStyle)
                                    }
                                    # As more output types come on, add logic here
                                }
                            }
                        }
                    }
                }
            }
        }
    }
}