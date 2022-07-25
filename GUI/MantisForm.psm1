
function Invoke-EventTracer {
    param (
        $ObjThis,
        $EventType
    )
    if ($Global:Verbose) {
        if (!$ObjThis.name) {
            Write-Host $ObjThis,$EventType -ForegroundColor Magenta
        } else {
            Write-Host $ObjThis.name,$EventType -ForegroundColor Magenta
        }
    }
    if (!$ObjThis.name) {
        Write-LogStep $ObjThis,$EventType ok
    } else {
        Write-LogStep $ObjThis.name,$EventType ok
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