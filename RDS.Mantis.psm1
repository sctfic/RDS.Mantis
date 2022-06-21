
#Get public and private function definition files.
$Public = @(Get-ChildItem -Path $PSScriptRoot\Public\*.ps1 -ErrorAction SilentlyContinue)
$Private = @(Get-ChildItem -Path $PSScriptRoot\Private\*.ps1 -ErrorAction SilentlyContinue)

#Dot source the files
Foreach ($import in @($Public + $Private)){
    try {
        . $import.fullname
    } catch {
        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
    }
}

# Create Aliases
New-Alias -Name Mantis -value Import-WFListBoxItem -Description "Load GUI of RDS.Mantis"

# Export all the functions Publics
Export-ModuleMember -Function $Public.Basename -Alias *