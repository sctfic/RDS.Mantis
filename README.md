# RDS.Mantis
GUI to manage RDS farm

## Run 
```powershell
    Import-module PSWinForm-Builder
    New-WinForm -DefinitionFile "$PSScriptRoot\GUI\MantisForm.psd1" -PreloadModules PsWrite,rds.mantis -Verbose
```
or
```powershell
    \RDS.Mantis\Mantis.ps1
```