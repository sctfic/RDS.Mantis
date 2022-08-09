
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
        $Ticks = 500,
        $ScriptBlock
    )
    begin {
        $MainTimer = New-Object System.Windows.Forms.Timer
        $MainTimer.Interval = $Ticks
        $Global:SequenceStart = New-Sequence @(
                {
                    $Global:ControlHandler['ListServers'].Enabled = $false
                    $Mantis.SelectedDomain.Servers.Get()
                },
                {
                    $Global:ControlHandler['TreeDFS'].Enabled = $false
                    $Mantis.SelectedDomain.DFS.Get()
                },
                {
                    $Global:ControlHandler['DataGridView_ADAccounts'].Enabled = $false
                    $Mantis.SelectedDomain.Users.Get()
                },
                {
                    $Global:ControlHandler['ListGroups'].Enabled = $false
                    $Mantis.SelectedDomain.Groups.Get()
                }
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
                            default {
                            }
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
        $Target.Enabled = $true
    }
}
function Convert-AdUser {
    <#
        .SYNOPSIS
            Converti des object ADSI et PSCustomObject
        .DESCRIPTION
            recupere les info de principale
        .PARAMETER Items
            
        .EXAMPLE
            Convert-AdsiUsers
        .NOTES
            Alban LOPEZ 2019
            alban.lopez@gmail.com
            http://git/PowerTech/
    #>
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline=$true)]$Items,
        [switch]$full,
        [switch]$DisplayPassword = $null
    )
    begin {
    }
    process {
        foreach ($item in $items) {
            try {
                $AdsiItem = [adsi]$item.DistinguishedName
                # write-host -ForegroundColor Green (Measure-Command {
                    try {
                        $rds = $AdsiItem.InvokeGet("AllowLogon")
                    } catch {
                        $rds = $null
                    } # n'as jamais été défini pour ce User
                    $TsAllowLogon = @()
                    if($rds -eq $true) {
                        $TsAllowLogon += 'TSE authorized!'
                    } elseif ($null -eq $rds) {
                        $TsAllowLogon += 'TSE not forbidden!'
                    }
                    if ($Item.mail -and $Item.msExchMailboxGuid) {
                        $TsAllowLogon += 'MailBox'
                    }
                    if (!$TsAllowLogon) {
                        $TsAllowLogon = 'Just Account!'
                    }
                    $TsAllowLogon =  $TsAllowLogon -join(' + ')

                #     if($AdsiItem.extensionattribute14 -match '\d{18}@.+\\.+'){
                #         ($lesspassTicks, $lesspassScope) = $AdsiItem.extensionattribute14 -split('@')
                #         $LesspassDate = (get-date ([int64]$lesspassTicks)).ToUniversalTime()
                #         if([Math]::Abs(($PwdLastSet - $LesspassDate).Totalseconds) -lt 120) {
                #             $PassWordType = $lesspassScope
                #         } else {
                #             $PassWordType = $null
                #         }
                #     } else {
                #         $PassWordType = $null
                #     }
                $Prop = [ordered]@{
                    NtAccountName        = $($Item.userPrincipalName) -replace ('(.+)@(.+)\..+', '$2\$1')
                    DisplayName          = "$($Item.DisplayName)"
                    Sid                  = "$($item.SID)"
                    Email                = "$($Item.mail)" # + $($AdsiItem.proxyAddresses)"
                    Type                 = $TsAllowLogon
                    Enabled                = "$($Item.Enabled)"
                    Phone                = "$($Item.telephonenumber)"
                    Prenom               = "$($Item.givenname)"
                    Nom                  = "$($Item.sn)"
                    Name                 = "$($Item.Name)"
                    State                = "$($Item.State)"
                    Tel_interne          = "$($Item.ipPhone)"
                    Password             = $null
                    Title                = "$($item.Title)"
                    AddressBookMembers   = "$($item.showInAddressBook)"
                    Description          = "$($Item.description)"
                    # PasswordType         = $PassWordType
                    PasswordDate         = $(try {$item.PasswordLastSet.ToString()} catch {''})
                    Groupes              = $($Item.memberOf | ?{$_} | %{(($_ -split(','))[0] -split('='))[1]}) # -join ("`n") # | ForEach-Object {([adsi]"LDAP://$_").name}
                    CodePostal           = "$($Item.postalcode)"
                    Ville                = "$($Item.City)"
                    SiteGeo              = "$($Item.State)"
                    Service              = "$($Item.Department)"
                    Bureau               = "$($Item.physicaldeliveryofficename)"
                    # OU                   = "$($AdsiItem.parent)"
                    Path                 = "$($Item.DistinguishedName)"
                    CreateDate           = $(try {$item.whenCreated.ToString()} catch {''})
                    Expire               = $(try {$item.AccountExpirationDate.ToString()} catch {''})
                    LastLogon            = $(try {$item.LastLogon.ToString()} catch {''})
                } | Write-Object -PassThru
                
                # write-verbose "lastlogontimestamp = [$($Item.Properties.lastlogontimestamp)]"
                # if($full){
                #     foreach($key in $item.Properties.GetEnumerator().name){
                #         # Write-Host $key,'-',$Prop.$Key,'-',$($item.Properties.$key) -fore red
                #         if(!$Prop.$Key -and $($item.Properties.$key)){
                #             try {
                #                 $Prop.$Key = $($item.Properties.$key)
                #             } catch {
                #                 Write-Error "[$key] Unreadabe !"
                #             }
                #         }
                #     }
                # }
            # }).TotalSeconds
                
                $Prop
            } catch {
                Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" '',$_ error
            }
        }
    }
    end {
        Write-Host 'End!'
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
        $Target.Visible = $true
    }
    process {
        foreach ($item in $Items) {
            $lastAdded = $Target.rows.Add(@($item.NtAccountName, $item.DisplayName, $item.Sid, $item.email, $item.Type, $item.State, $item.Phone, $item.Description, $item.Facturation, $item.PasswordType, $item.CreateDate, $item.LastLogonDate, $item.Expire.ToString(), ($item.Groupes -join (", "))))
            $lastRow = $Target.rows[$lastAdded]

            if ($item.State -eq 'Disabled') {
                $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DimGray # ColorComptesInActif
            } elseif ($item.Type -match 'Session') {
                $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::MediumBlue # ColorComptesActif
            } elseif ($item.Type -match 'Bal') {
                $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::SlateBlue # ColorBalOnly
            } else {
                $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::DarkViolet # ColorComptesSystem
            }

            if ($item.ExpireDate -and ($item.ExpireDate).AddDays(300) -gt (Get-Date)) {
                $lastRow.DefaultCellStyle.ForeColor = [system.Drawing.Color]::White
                $lastRow.DefaultCellStyle.BackColor = [system.Drawing.Color]::Tomato
            }
        }
    }
    end {
        $Target.ResumeLayout()
        $Target.Enabled = $true
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
        $Target.Visible = $true
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
        $Target.Enabled = $true
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