
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

function Start-WatchTimer {
    [CmdletBinding()]
    param (
        $Ticks = 500
    )
    begin {
        $MainTimer = New-Object System.Windows.Forms.Timer
        $MainTimer.Interval = $Ticks
        $Global:SequenceStart = New-Sequence @(
                # {$Mantis.SelectedDomain.Servers.Get()},
                # {$Mantis.SelectedDomain.DFS.Get()},
                {$Mantis.SelectedDomain.Users.Get()}
                # {$Mantis.SelectedDomain.Groups.Get()}
                # {$Mantis.SelectedDomain.Trash.Get()}
            )
    }
    process {
        $MainTimer.add_Tick({ # on supprime les jobs terminé, toute les seconde
            Import-Module PsWrite
            $Threads = Get-Job | ?{$_.Name -like 'Mantis_*' -and $_.State -eq 'Completed'}
            if ($Threads) {
                foreach($thread in $Threads) {
                    try {
                        switch -Regex ($thread.Name) {
                            'Mantis_Srv_.+' {
                                $Mantis.SelectedDomain.Servers.Get() | Update-MantisSrv
                            }
                            'Mantis_Sessions_.+' {
                                $Mantis.SelectedDomain.Servers.GetRDSessions() | Update-MantisSessions
                            }
                            'Mantis_Usr_.+' {
                                $Mantis.SelectedDomain.Users.Get() | Update-MantisUsers
                            }
                            'Mantis_Grp_.+' {
                                $Mantis.SelectedDomain.Groups.Get() | Update-MantisGrp
                            }
                            'Mantis_Trash_.+' {
                                $Mantis.SelectedDomain.Trash.Get() | Update-MantisTrash
                            }
                            # 'Mantis_DFS_.+' { # pas de thread, tres rapide en mode AD
                            #     Write-Host 'DFS ici' -BackgroundColor Magenta
                            #     $Mantis.SelectedDomain.DFS | Update-MantisDFS
                            #     Write-Host 'DFS la' -BackgroundColor Magenta
                            # }
                            default {}
                        }
                        Write-Host $thread.Name -ForegroundColor DarkGreen
                    } catch {
                        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                    }
                }
            } elseif(!(Get-Job | ?{$_.Name -like 'Mantis_*'})) {
                Get-Job | Select-Object Name,Id,State | Write-Object -fore DarkRed
                $next = $Global:SequenceStart.Next()
                if ($Next) {Invoke-Command $next}
            }

            $ThreadsInProgress = Get-Job | ?{$_.Name -like 'Mantis_*' -and $_.State -notlike 'Completed'}
            if ($ThreadsInProgress) {
                foreach($thread in $ThreadsInProgress) {
                    try {
                        switch -Regex ($thread.Name) {
                            'Mantis_Srv_.+' { $WinControl = 'ListServers' }
                            'Mantis_Sessions_.+' { $WinControl = 'DataGridView_Sessions' }
                            # 'Mantis_DFS_.+' { $WinControl = 'TreeDFS' } # pas de thread, tres rapide en mode AD
                            'Mantis_Usr_.+' { $WinControl = 'DataGridView_ADAccounts' }
                            'Mantis_Grp_.+' { $WinControl = 'ListGroups' }
                            'Mantis_Trash_.+' { $WinControl = 'ListTrash' }
                            default {$WinControl = ($thread.Name -split('_'))[1]}
                        }
                        Start-Loading $WinControl $thread
                    } catch {
                        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                    }
                }
            }
        })
    }
    end {
        $MainTimer.Start() | Out-Null
    }
}
function Start-Loading {
    param (
        $WinControl,
        $Thread
    )
    try {
        # $Global:ControlHandler[$WinControl].Enabled = $false
        $Global:ControlHandler["ProgressBar_$WinControl"].Visible = $true
        # $thread | Select-Object Name,Id,State | Write-Object -fore DarkMagenta
    } catch {
        # Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
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
        $lst = $Items | ForEach-Object {
                if($_.RDS.MemberOfFarm) {
                    $Grp = $_.RDS.MemberOfFarm
                } elseif ($_.RDS.ServerBroker) {
                    $Grp = 'Brocker'
                } elseif($_.isDC) {
                    $Grp = 'Domain Controler'
                } else {
                    $Grp = 'Others'
                }
            [PSCustomObject]@{
                FirstColValue = $_.Name
                NextValues = @($_.IP,$_.OperatingSystem,$_.InstallDate,$_.RDS.ProductVersion)
                Group   = $Grp
                Caption = $(if($_.RDS.ServerBroker){"Broker [$($_.RDS.ServerBroker)]"} else {''})
                Status  = $Null
                Shadow  = (!$_.OperatingSystem)
                Tag = $_.DistinguishedName
            }
        }
        
        $lst | Where-Object {
            $_.Group -ne 'Others'
        } | Update-ListView -listView $Target
        $lst | Where-Object {
            $_.Group -eq 'Others'
        } | Update-ListView -listView $Target
    }
    end {
        $Target.EndUpdate()
        $Target.Enabled = $true
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            # Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen

        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
    }
}
function Update-MantisTrash {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['ListTrash']
    )
    begin {
        $Target.BeginUpdate()
        $Target.Items.Clear()
    }
    process {
        $Items | ForEach-Object {
            try {
                $fullName = $_.dNSHostName
            } catch {
                try {
                    $fullName = $_.sAMAccountName
                } catch {}
            }
            [PSCustomObject]@{
                FirstColValue = $_.'msDS-LastKnownRDN'
                NextValues = @($FullName, $_.whenChanged, $_.LastKnownParent)
                Group   = $_.ObjectClass
                Caption = "SID: $($_.objectSid)`nDescription: $($_.Description)"
                Status  = ''
                Shadow  = $false
                Tag = $_.DistinguishedName
            } | Update-ListView -listView $Target
        }
    }
    end {
        $Target.EndUpdate()
        $Target.Enabled = $true
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            # Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
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
                NextValues = @($_.Members.count, $_.MemberOf)
                Group   = if($_.GroupCategory){$_.GroupCategory}else{'Distribution'}
                Caption = $(($_.Members | ?{$_} | %{(($_ -split(','))[0] -split('='))[1]}) -join("`n"))
                Status  = ''
                Shadow  = (!$_.Members.count)
                Tag = $_.DistinguishedName
            } | Update-ListView -listView $Target
        }
    }
    end {
        $Target.EndUpdate()
        $Target.Enabled = $true
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            # Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen

        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
    }
}
function Update-MantisUsers {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['DataGridView_ADAccounts']
    )
    begin {
        $Target.SuspendLayout()
        $Target.Rows.Clear()
        # $Target.Enabled = $false
        $Target.Visible = $true
    }
    process {
        foreach ($item in $Items) {
            try {
                $lastAdded = $Target.rows.Add(@($item.NtAccountName, $item.DisplayName, $item.email, $item.Type, $item.Status, $item.Phone, $item.Description, $item.OU, $item.Sid, $item.PasswordType, $item.CreateDate, $item.LastLogonDate, $item.Expire, ($item.Groupes -join ("`n"))))
                $lastAdded.Tag = $item.DistinguishedName
                $lastRow = $Target.rows[$lastAdded]
                
                if ($item.Status -eq 'Disabled') {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DimGray # ColorComptesInActif
                } elseif ($item.Type -match 'TSE.+') {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::MediumBlue # ColorComptesActif
                } elseif ($item.Type -match 'MailBox') {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::SlateBlue # ColorBalOnly
                } else {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DarkViolet # ColorComptesSystem
                }
                
                if ($item.ExpireDate -and ($item.ExpireDate).AddDays(300) -gt (Get-Date)) {
                    $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::White
                    $lastRow.DefaultCellStyle.BackColor = [system.Drawing.Color]::Tomato
                }
            } catch {
                Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
            }
        }
    }
    end {
        $Target.ResumeLayout()
        $Target.Enabled = $true
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen

        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
    }
}
function Update-MantisSessions {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]$Items,
        $Target = $Global:ControlHandler['DataGridView_Sessions']
    )
    begin {
        # $Target.Enabled = $false
        $Target.Visible = $true
        $Target.SuspendLayout()
        $Target.Rows.Clear()
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
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            # Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen

        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
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
        $Target.Enabled = $true
        try {
            $Global:ControlHandler["ProgressBar_$($Target.name)"].Visible = $false
            Write-Host "ProgressBar_$($Target.name)" -ForegroundColor DarkGreen

        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "Missing ProgressBar","ProgressBar_$($Target.name)"  error
        }
    }
}
function Convert-DGV_RDS_Row {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]$Rows = $Global:ControlHandler['DataGridView_Sessions'].SelectedRows 
    )
    begin {
        
    }
    process {
        foreach ($Row in $Rows) {
            # Write-Object $Row.cells
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
function Set-SelectedRDServers {
    param($ListView = $Global:ControlHandler['ListServers'])
    $Global:LVSrvChange = $False
    $Selected = $ListView.SelectedItems.Text
    $Computers = $Global:mantis.SelectedDomain.Servers.items | Where-Object {
        $Selected -contains $_.Name
    }
    $Global:mantis.SelectedDomain.Servers.GetRDSessions($Computers)
}
function Stop-DGV_RDSSessions {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName = $true)] $ComputerName = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $SessionID = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $NtAccountName = $null
    )
    begin {
        $Listing = @()
        $CassiaSession = @()
    }
    process {
        # $ComputerName, $SessionID, $NtAccountName | Write-Object
        $CassiaSession += $ComputerName | Get-RdComputer | Get-RdSession -Identity $SessionID # | ?{$NtAccountName}
    }
    end {
        $Listing = $CassiaSession | ForEach-Object {
            "`t$($_.server.servername):$($_.SessionId) -> $($_.UserAccount)"
        }
        # Write-Object $CassiaSession -foreGroundColor DarkBlue
        if ([System.Windows.Forms.MessageBox]::Show(
                "Fermeture de $($CassiaSession.Count) session(s) :`n`n$($Listing -join("`n"))",
                "Question", [System.Windows.Forms.MessageBoxButtons]::YesNo) -eq 'Yes'
        ) {
            foreach ($Target in $CassiaSession) {
                try {
                    $Target.Logoff($false)
                    Write-LogStep "Fermeture de session [$($Target.Server.ServerName):$($Target.SessionId)]", "$([string]$Target.UserAccount)" OK
                }
                catch {
                    Write-LogStep "Fermeture de session [$($Target.Server.ServerName):$($Target.SessionId)]", $_ Error
                }
            }
            $true
        }
    }
}



























if (!(Get-Module PsWrite)) {
    function Script:Write-logstep {
        param ( [string[]]$messages, $mode, $MaxWidth, $EachLength, $prefixe, $logTrace )
        Write-Verbose "$($messages -join(',')) [$mode]"
    }
}
Write-LogStep 'Chargement du module ', $PSCommandPath warn