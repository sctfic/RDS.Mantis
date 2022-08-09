param([switch]$verbose=$null)
    Import-module PSWinForm-Builder
    New-WinForm -DefinitionFile "$PSScriptRoot\GUI\MantisForm.psd1" -Verbose # -PreloadModules PsWrite,rds.mantis # -Verbose:$verbose