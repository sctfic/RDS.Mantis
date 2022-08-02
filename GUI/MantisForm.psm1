
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
function Start-WatchTimer {
    [CmdletBinding()]
    param (
        $Ticks = 500,
        $ScriptBlock
    )
    begin {
        $MainTimer = New-Object System.Windows.Forms.Timer
        $MainTimer.Interval = $Ticks
        $Global:SequenceStart = New-Sequence @(
                {$Mantis.CurrentDomain.Servers.Get()},
                {$Mantis.CurrentDomain.DFS.Get()},
                {$Mantis.CurrentDomain.Sessions.Get()},
                {$Mantis.CurrentDomain.Users.Get()},
                {$Mantis.CurrentDomain.Groups.Get()}
            )
    }
    process {
        $MainTimer.add_Tick({ # on supprime les jobs terminé, toute les seconde
            $Threads = Get-Job | ?{$_.Name -like 'Mantis_*' -and $_.State -eq 'Completed'}
            if ($Threads) {
                
                foreach($thread in $Threads) {
                    # Write-Host ' > > > ',$thread.id,$thread.Name,$thread.State -ForegroundColor red
                    try {
                        switch -Regex ($thread.Name) {
                            'Mantis_Srv_.+' {
                                $Mantis.SelectedDomain.Servers.Get() | Update-MantisSrv
                            }
                            'Mantis_Usr_.+' {
                                $Mantis.SelectedDomain.Users.Get() | Update-MantisUsr
                            }
                            'Mantis_Grp_.+' {
                                $Mantis.SelectedDomain.Groups.Get() | Update-MantisGrp
                            }
                            'Mantis_Sessions_.+' {
                                $Mantis.SelectedDomain.Sessions.Get() | Update-MantisSessions
                            }
                            'Mantis_DFS_.+' {
                                $Mantis.SelectedDomain.DFS.Get() | Update-MantisDFS
                            }
                            default {}
                        }
                    } catch {
                        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                    }
                }
            } elseif (!(Get-Job | Where-Object{$_.Name -like 'Mantis_*' -and $_.State -ne 'Completed'})) {
                $next = $Global:SequenceStart.Next()
                if ($Next) {Invoke-Command $next}
            }
        })
        $MainTimer.Start() | Out-Null
    }
    end {
        
    }
}

function Update-MantisSrv {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['ListServers']
    )
    begin {
        $Target.BeginUpdate()
        $Target.Items.Clear()
    }
    process {
        $Items | ForEach-Object {
            [PSCustomObject]@{
                FirstColValue = $_.Name
                NextValues = @($_.IP,$_.OperatingSystem,$_.DN)
                Group   = $null
                Caption = $null
                Status  = ''
                Shadow  = (!$_.IP)
            } | Update-ListView -listView $Target
        }
    }
    end {
        $Target.EndUpdate()
    }
}
function Update-MantisUsr {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['DataGridView_ADAccounts']
    )
    begin {
        # $Target.BeginUpdate()
        # $Target.Items.Clear()
    }
    process {
        $Items | ForEach-Object {
            [PSCustomObject]@{

            }
        }
    }
    end {
        # $Target.EndUpdate()
    }
}
function Update-MantisGrp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['ListGroups']
    )
    begin {
        $Target.BeginUpdate()
        $Target.Items.Clear()
    }
    process {
        $Items | ForEach-Object {
            [PSCustomObject]@{
                FirstColValue = $_.Name
                NextValues = @($_.Members.count, $_.MemberOf, $_.DistinguishedName)
                Group   = if($_.GroupCategory){$_.GroupCategory}else{'Distribution'}
                Caption = $null
                Status  = ''
                Shadow  = (!$_.Members.count)
            } | Update-ListView -listView $Target
        }
    }
    end {
        $Target.EndUpdate()
    }
}
function Update-MantisSessions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['']
    )
    begin {
        $Target.BeginUpdate()
        $Target.Items.Clear()
    }
    process {
        $Items | ForEach-Object {
            [PSCustomObject]@{

            } | Update-ListView -listView $Target
        }
    }
    end {
        $Target.EndUpdate()
    }
}
function Update-MantisDFS {
    [CmdletBinding()]
    param (
        $Dfs = $mantis.SelectedDomain.DFS,
        $Target = $Global:ControlHandler['TreeDFS']
    )
    begin {
        $Target.Nodes.Clear()
        # $Target.BeginUpdate()
        [PSCustomObject]@{
            Name =  $Dfs.Root
            Handler = $Dfs.Root
            ToolTipText = $Dfs.Root
        } | Update-TreeView -treeNode $Target -Expand -Depth 1 -ChildrenScriptBlock {[PSCustomObject]@{
            Name =  'Empty (No DFS Found!)'
            Handler = ''
            ToolTipText = 'Impossible de trouver une racine DFS'
            ForeColor = [System.Drawing.Color]::Gray
        }}
    }
    process {
        try {
            $DFSNodes = $Dfs.Get() | Where-Object { $_.Name }
            if ($DFSNodes) {
                $DFSNodes | ForEach-Object {
                    [PSCustomObject]@{
                        Name =  $_.Name
                        Handler = $_.Path
                        ToolTipText = $_.Path
                        ForeColor = [System.Drawing.Color]::DarkBlue
                    }
                } | Update-TreeView -treeNode $Target.TopNode -Clear -ChildrenScriptBlock {
                    param($item)
                    $item.Handler | get-childitem -Directory -Force -ea 0 | ForEach-Object {
                        @{
                            Name = $_.Name
                            Handler = $_.FullName
                            ToolTipText = "$Prefix$($_.FullName)"
                            ForeColor = [System.Drawing.Color]::DarkCyan
                        }
                    }
                } # -Expand <!> declenche l'event d'expand
            }
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }

    }
    end {
        $Target.EndUpdate()
    }
}