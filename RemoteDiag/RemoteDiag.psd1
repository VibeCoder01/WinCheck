@{
    RootModule = 'RemoteDiag.psm1'
    ModuleVersion = '0.1.0'
    GUID = 'f4b5ac7b-6156-4f78-a5c8-f4cbf4f8ce44'
    Author = 'WinCheck'
    CompanyName = 'WinCheck'
    Copyright = '(c) WinCheck'
    Description = 'Remote diagnostics snapshot and performance tooling for Windows support triage.'
    PowerShellVersion = '7.0'
    FunctionsToExport = @('Get-RDHostSnapshot', 'Get-RDPerformanceSample', 'Export-RDReport', 'Compare-RDHostSnapshot')
    CmdletsToExport = @()
    VariablesToExport = @()
    AliasesToExport = @()
}
