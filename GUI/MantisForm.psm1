
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
                            'Mantis_Usr_.+' {
                                $Mantis.SelectedDomain.Users.Get() | Update-MantisUsr
                            }
                            'Mantis_Grp_.+' {
                                $Mantis.SelectedDomain.Groups.Get() | Update-MantisGrp
                            }
                            'Mantis_Session_.+' {
                                $Mantis.SelectedDomain.Session.Get() | Update-MantisSessions
                            }
                            default {}
                        }
                    } catch {
                        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                    }
                }
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
                NextValues = @($_.Members.count, $_.DN)
                Group   = $_.GroupCategory
                Caption = $null
                Status  = ''
                Shadow  = $null
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
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items = $mantis.SelectedDomain.DFS,
        $Target = $Global:ControlHandler['TreeDFS']
    )
    begin {
        $Target.BeginUpdate()
        $Target.Items.Clear()
    }
    process {


        $Items.Get() | ForEach-Object {
            [PSCustomObject]@{
                Name =  $_.Name
                Handler = $_.Path
                ToolTipText = $_.Path
                ForeColor = [System.Drawing.Color]::DarkBlue
            }
        } | Write-Object -PassThru -fore Yellow | Update-TreeView -treeNode $Target.TopNode 
        # -expand -Depth 2 -ChildrenScriptBlock {
        #     param($parent)
        #     $parent | Write-Object -PassThru -fore DarkMagenta
        #     $parent.Handler | Write-Object -PassThru -fore Cyan | Get-ChildItem | ForEach-Object {
        #         [PSCustomObject]@{
        #             Name =  $_.Name
        #             Handler = $_.FullPath
        #             ToolTipText = $_.FullPath
        #             ForeColor = [System.Drawing.Color]::Blue
        #         }
        #     }
        # }
    }
    end {
        $Target.EndUpdate()
    }
}