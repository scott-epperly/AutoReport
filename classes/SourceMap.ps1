class SourceMap {
    [string]$DataProviderName;
    [string]$QueryFileName;
    [string]$DataTableId;
}

Class SourceMapExcel : SourceMap {
    [string]$WorksheetName;
    [string]$Cell;
    [string]$TableStyle;
}

Class SourceMapDelimited : SourceMap {
    [string]$Delimiter;
    [string]$Encoding;
}