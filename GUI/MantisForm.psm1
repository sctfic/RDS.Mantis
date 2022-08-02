
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
                # {$Mantis.CurrentDomain.DFS.Get()},
                # {$Mantis.CurrentDomain.Users.Get()},
                # {$Mantis.CurrentDomain.Groups.Get()}
                {}
            )
    }
    process {
        $MainTimer.add_Tick({ # on supprime les jobs terminé, toute les seconde
            $Threads = Get-Job | ?{$_.Name -like 'Mantis_*' -and $_.State -eq 'Completed'}
            if ($Threads) {
                
                foreach($thread in $Threads) {
                    try {
                        switch -Regex ($thread.Name) {
                            'Mantis_Srv_.+' {
                                $Mantis.SelectedDomain.Servers.Get() | Update-MantisSrv
                            }
                            'Mantis_Sessions_.+' {
                                Write-Host "Mantis.SelectedDomain.Servers.GetRDSessions()" -fore DarkMagenta
                                $Mantis.SelectedDomain.Servers.GetRDSessions() | Update-MantisSessions
                            }
                            'Mantis_Usr_.+' {
                                $Mantis.SelectedDomain.Users.Get() | Update-MantisUsr
                            }
                            'Mantis_Grp_.+' {
                                $Mantis.SelectedDomain.Groups.Get() | Update-MantisGrp
                            }
                            'Mantis_DFS_.+' {
                                Write-Center 'DFS ici'
                                $Mantis.SelectedDomain.DFS | Update-MantisDFS
                                Write-Center 'DFS la'

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
        $Target.Groups.Clear()
    }
    process {
        # $Target.Groups.AddRange(@($items.RDS.MemberOfFarm | ForEach-Object { $_ } | Sort-Object -Unique))
        # $Target.Groups.Add('Domain Controler')
        # $Target.Groups.Add('Brocker')
        # $Target.Groups.Add('Other')
        $Items | ForEach-Object {
            [PSCustomObject]@{
                FirstColValue = $_.Name
                NextValues = @($_.IP,$_.OperatingSystem,$_.InstallDate,$_.RDS.ProductVersion,$_.DN)
                Group   = $(if($_.RDS.MemberOfFarm) {$_.RDS.MemberOfFarm} elseif ($_.RDS.ServerBroker) {'Brocker'} elseif($_.isDC) {'Domain Controler'} else {'Other'})
                Caption = $(if($_.RDS.ServerBroker){"Broker [$($_.RDS.ServerBroker)]"} else {''})
                Status  = ''
                Shadow  = (!$_.IP)
            }
        } | Update-ListView -listView $Target
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
        $Target.SuspendLayout()
        $Target.Rows.Clear()
        $Target.Enabled = $false
    }
    process {
        foreach ($item in $Items) {
            
        }
    }
    end {
        $Target.ResumeLayout()
        $Target.Enabled = $true
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
        [Parameter(ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['DataGridView_Sessions']
    )
    begin {
        $Target.SuspendLayout()
        $Target.Rows.Clear()
        $Target.Enabled = $false
    }
    process {
        foreach ($item in ($Items | ?{$_})) {
            try {
                $lastAdded = $Target.rows.Add(@("$($item.ComputerName)", "$($item.SessionID)", "$($item.State)", "$($item.Sid)", "$($item.NtAccountName)", "$($item.IPAddress)", "$($item.ClientName)", "$($item.Protocole)", "$($item.ClientBuildNumber)", "$($item.LoginTime)", "$($item.DisconnectTime)", "$($item.ConnectTime)", "$($item.Inactivite)", "$($item.Screen)", "$($item.Process)"))
                $lastRow = $Target.rows[$lastAdded]
                if ($item.State -like "HS") {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::White
                    $lastRow.DefaultCellStyle.BackColor = [system.Drawing.Color]::DarkGray
                } elseif ($item.State -like "Linux") {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::Purple
                    $lastRow.DefaultCellStyle.BackColor = [system.Drawing.Color]::SandyBrown
                } elseif ($item.State -Like "Active") {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DarkBlue
                } elseif ($item.Protocole -eq 'Service' -or $item.State -eq 'Listening') {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DimGray
                } else {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::CornflowerBlue
                }
            } catch {
                Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
            }
        }
    }
    end {
        $Target.ResumeLayout()
        $Target.Enabled = $true
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
                    [PSCustomObject]@{
                        Name = '-'
                        Handler = '-'
                        ToolTipText = 'Impossible de trouver une racine DFS'
                        ForeColor =[system.Drawing.Color]::LightGray
                    }
                    # param($item)
                    # $item.Handler | get-childitem -Directory -Force -ea 0 | ForEach-Object {
                    #     @{
                    #         Name = $_.Name
                    #         Handler = $_.FullName
                    #         ToolTipText = "$Prefix$($_.FullName)"
                    #         ForeColor = [System.Drawing.Color]::DarkCyan
                    #     }
                    # }
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

function Convert-DGVRow {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Rows
    )
    begin {
        
    }
    process {
        foreach ($Row in $Rows) {
            Write-Object $Row.cells
            [pscustomobject][ordered]@{
                ComputerName  = $($Row.cells[0].value)
                SessionID     = $( try { [int]$Row.cells[1].value } catch { $null } )
                State         = $($Row.cells[2].value)
                Sid           = $(if ($Row.cells[3].value -like 'S-1-*') {$sid = $Row.cells[3].value} else {$sid = $null})
                NtAccountName = $($Row.cells[4].value)
                IPAddress     = $($Row.cells[5].value)
                ClientName    = $($Row.cells[6].value)
                Protocole     = $($Row.cells[7].value)
                handler       = $(if ($handler) { $Row }else { $null }) # la ligne dans le DataGridView
            }
        }
    }
    end {
        
    }
}
function Start-ActionByRow {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Rows,
        $ColumnClicked = $null
    )
    begin {
        
    }
    process {
        foreach ($Row in $Rows) {
            Invoke-EventTracer 'Start-ActionByRow' $ColumnClicked
            $row | Write-Object -PassThru
            switch ($ColumnClicked) {
                'ComputerName' {}
                'UserAccount' {}
                'IpAddress' {}
                default {}
            }
        }
    }
    end {
        
    }
}