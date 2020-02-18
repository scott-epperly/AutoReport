class Query{
    [string]$DataProviderName;
    [string]$QueryFileName;
    [string]$QueryText;
    [string[]]$QueryParameters;
    [int]$MinExecutionOrder;
    [int]$MaxExecutionOrder;

    #--------------#
    # Constructors #
    #--------------#
    Query([System.IO.FileInfo]$QueryFile, [string]$DataProviderName, [int]$MinExecutionOrder, [int]$MaxExecutionOrder){
        #Read the file contents into QueryText property
        $this.DataProviderName = $DataProviderName;
        $this.QueryFileName = $QueryFile.Name;
        $this.MinExecutionOrder = $MinExecutionOrder;
        $this.MaxExecutionOrder = $MaxExecutionOrder;
        $this.QueryText = [System.IO.File]::ReadAllText($QueryFile.FullName);
        
        #TODO: Build out getting a list of query parameters for all other Data Providers
        if ($DataProviderName -eq "Kusto") {
            $this.QueryParameters = Get-PSAdxCSLParameter -Path $QueryFile.FullName;
        }
    }
    
    #---------#
    # Methods #
    #---------#

    # Query SQL Server
    [object[]]ExecuteSQL ([string]$TargetConnectionString, [string]$DataProviderName, [object]$UserParamHash){
        
        # load result into a DataTable
        [System.Data.DataSet]$ds = [System.Data.DataSet]::new();

        # Build appropriate objects based on $DataProviderName
        switch($DataProviderName) {
            "SqlClient" {
                Write-Debug "[Query.ExecuteSQL] Executing SqlClient"
                [System.Data.SqlClient.SqlConnection]$DBConnection = [System.Data.SqlClient.SqlConnection]::new($TargetConnectionString);
                $DBConnection.Open()

                $sqlcmd = [System.Data.SqlClient.SqlCommand]::new();
                $sqlcmd.Connection = $DBConnection;
                $sqlcmd.CommandText = $this.QueryText;

                # Bind Parameters
                foreach ($key in $UserParamHash.Keys) {
                    $sqlcmd.Parameters.AddWithValue($key, $UserParamHash[$key]);
                }
                
                [System.Data.SqlClient.SqlDataAdapter]$da = [System.Data.SqlClient.SqlDataAdapter]::new($sqlcmd);
                $da.Fill($ds);
                #$ds.Load($sqlcmd.ExecuteReader()) | Out-Null;

                $DBConnection.Close()
            }
            default {
                #attempt SqlClient by default
                Write-Error "Invalid DataProviderName: $DataProviderName"
            }
        }
        
        return $ds.Tables;
    }

    # Query Kusto (Azure Data Explorer)
    [object[]]ExecuteKQL ([string]$TargetConnectionString, [object]$UserParamHash){
        Write-Debug "[Query.ExecuteKQL()] TargetConnectionString: $TargetConnectionString";
        Write-Debug "[Query.ExecuteKQL()] UserParamHash: $UserParamHash";
        Write-Debug "[Query.ExecuteKQL()] QueryText: $($this.QueryText)";

        return Invoke-PSAdxQuery -ConnectionString $TargetConnectionString -DatabaseName ($TargetConnectionString | Select-String -Pattern "Catalog\=(\w*)").Matches.Groups[1].Value -Query $this.QueryText -QueryParameters $UserParamHash;
    }

    [System.Collections.ArrayList]GetParameter() {
        [System.Collections.ArrayList]$foundParams = [System.Collections.ArrayList]::new();
        try {
            $regMatches = (($this.QueryText | Select-String "^declare query_parameters.*").Matches[0].Value | Select-String -Pattern "(\w*):\w*" -AllMatches -ErrorAction SilentlyContinue)
            foreach ($item in $regMatches.Matches.Groups)
            {
                if ($item.Value -notmatch "(\(|\)|\,|\:|\^|\"")") 
                {
                    $foundParams.Add([string]($item.Value).Trim()) | Out-Null   
                }
            }
        }
        catch
        {}
        return $foundParams
    }
}