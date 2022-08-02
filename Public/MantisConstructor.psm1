function Invoke-MantisConstructor {
    [CmdletBinding()]
    param ()
    begin {
        try {
            $items = @()
            $Global:mantis = New-Mantis
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }
        $ChildrenScriptBlock = {
            param($parent)
            Get-ADOrganizationalUnit -Server $parent.Server -SearchBase $parent.Handler -Filter * -SearchScope OneLevel -ea SilentlyContinue | %{
                @{
                    Name = $_.Name
                    Server = $parent.Server
                    Handler = $_.DistinguishedName
                    ToolTipText = $_.DistinguishedName
                    ForeColor = [System.Drawing.Color]::DarkBlue # ::FromArgb(64,64,64)
                }
            }
        }
    }
    process {
        try {
            $mantis.CurrentDomain | ForEach-Object {
                @{
                    Name = $_.Name
                    Server = $_.Name
                    Handler = $_.DistinguishedName
                    ToolTipText = $_.DistinguishedName
                }
            } | Update-TreeView -treeNode $Global:ControlHandler['TreeForest'] -Clear -ChildrenScriptBlock $ChildrenScriptBlock -Depth 2
            $mantis.TrustedDomain | ForEach-Object {
                @{
                    Name = $_.Name
                    Server = $_.Name
                    Handler = $_.DistinguishedName
                    ToolTipText = $_.DistinguishedName
                    ForeColor = [system.Drawing.Color]::DarkRed
                }
            } | Update-TreeView -treeNode $Global:ControlHandler['TreeForest'] -ChildrenScriptBlock $ChildrenScriptBlock -Depth 0
            # $mantis.Domain('pep64.org').Servers.get()
            
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }
    }
    end {
        Start-WatchTimer -ticks 250
    }
}








if (Get-Module PsWrite) {
    # Export-ModuleMember -Function Convert-RdSession, Get-RdSession
    Write-LogStep 'Chargement du module ', $PSCommandPath ok
} else {
    function Script:Write-logstep {
        param ( [string[]]$messages, $mode, $MaxWidth, $EachLength, $prefixe, $logTrace )
        Write-Verbose "$($messages -join(',')) [$mode]"
        # Write-LogStep '',"" ok
    }
}