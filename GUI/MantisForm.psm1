
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
                {$Mantis.SelectedDomain.Servers.Get()},
                {$Mantis.SelectedDomain.DFS.Get()},
                {$Mantis.SelectedDomain.Groups.Get()},
                {$Mantis.SelectedDomain.Trash.Get()},
                {$Mantis.SelectedDomain.Users.Get()}
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
                            'Mantis_AskAQuestionToUser_.+' {
                                Receive-Job $thread.Name -Wait -AutoRemoveJob | Update-LogsOfRequest -target $Global:ControlHandler['RichTextBox_Logs']
                            }
                            default {
                                $TargetName = ($thread.Name -split('_'))[1]
                                try {
                                    $Target = $Global:ControlHandler[$TargetName]
                                    # $Target.BeginUpdate()
                                    # $Target.Items.Clear()
                                } catch {
                                    Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                                }
                                Receive-Job $thread.Name -Wait -AutoRemoveJob | Update-ListView -listView $Target
                                try {
                                    # $Target.Enabled = $true
                                    $Global:ControlHandler["ProgressBar_$TargetName"].Visible = $false
                                } catch {
                                    Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                                }
                            }
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
function Update-LogsOfRequest {
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$SessionAnswer,
        $Target = $Global:ControlHandler['RichTextBox_Logs']
    )
    begin{
    # [PSCustomObject]@{
    #     Request         = $SessionAnswer.Request
    #     ComputerName    = $SessionAnswer.Server.ServerName
    #     UserAccount     = $SessionAnswer.UserAccount
    #     Answer          = $SessionAnswer.Answer
    #     ClientName      = $SessionAnswer.ClientName
    #     ClientIPAddress = $SessionAnswer.ClientIPAddress
    # }
    }
    process {
        if ("$($SessionAnswer.Answer)" -eq 'no') {
            $color = 'Red'
            $Answer = '   NO   '
        } else {
            $color = 'Green'
            $Answer = '  YES   '
        }
        $Target.SelectionColor = 'black'
        $Target.AppendText('[')
        $Target.SelectionColor = $color
        $Target.AppendText($Answer)
        $Target.SelectionColor = 'black'
        $Target.AppendText("]`t-> Reponce de ")
        $Target.SelectionColor = 'Blue'
        $Target.AppendText($SessionAnswer.UserAccount)
        $Target.SelectionColor = 'Black'
        $Target.AppendText(' depuis ')
        $Target.SelectionColor = 'Blue'
        $Target.AppendText($SessionAnswer.ComputerName)
        $Target.SelectionColor = 'Gray'
        $Target.AppendText(' a la Question: ')
        $Target.SelectionColor = $color
        $Target.AppendText($SessionAnswer.Request)
        $Target.SelectionColor = 'DarkGray'
        $Target.AppendText(" $(get-date)`r`n")
        $Target.ScrollToCaret()
    }
    End {}
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
function Start-Job4RightProperties {

    if ($Global:ControlHandler['UserNameGbx'].Text -ne $Global:RDS_LastSelected.NtAccountName) {
        $Job = Get-Job -Name "Mantis_LVUserADProp_$($mantis.SelectedDomain.Name)" -ea 0
        if ((!$Job -or $job.PSBeginTime -lt (Get-Date).AddSeconds(-10))){
            $Job | Remove-Job -Force
            $Global:ControlHandler['UserNameGbx'].Text = $Global:RDS_LastSelected.NtAccountName
            $Global:ControlHandler['LVUserADProp'].Enabled = $false
            $Global:ControlHandler['LVUserADProp'].Items.Clear()

            Start-ThreadJob `
                -Name "Mantis_LVUserADProp_$($mantis.SelectedDomain.Name)" `
                -InitializationScript {} `
                -ScriptBlock {
                    param($item)
                    Import-Module ActiveDirectory,PsWrite,RDS.Mantis
                    Import-module PSBright -Function Get-MailInfos,Get-MailProvider,Get-ExchMailProperty -SkipEditionCheck -DisableNameChecking
                    # $item | Write-Object -PassThru
                    $item | Update-MantisUserProp
                } -ArgumentList $Global:RDS_LastSelected
        }

        $Job = Get-Job -Name "Mantis_LVUserGroups_$($mantis.SelectedDomain.Name)" -ea 0
        if ((!$Job -or $job.PSBeginTime -lt (Get-Date).AddSeconds(-10))){
            $Job | Remove-Job -Force
            $Global:ControlHandler['UserNameGbx'].Text = $Global:RDS_LastSelected.NtAccountName
            $Global:ControlHandler['LVUserGroups'].Enabled = $false
            $Global:ControlHandler['LVUserGroups'].Items.Clear()
    
            Start-ThreadJob `
                -Name "Mantis_LVUserGroups_$($mantis.SelectedDomain.Name)" `
                -InitializationScript {} `
                -ScriptBlock {
                    param($item)
                    Import-Module ActiveDirectory,PsWrite,RDS.Mantis
                    # $item | Write-Object -PassThru
                    $item | Update-MantisUserGroups
                } -ArgumentList $Global:RDS_LastSelected
        }
    }

    if ($Global:ControlHandler['ServerNameGbx'].Text -ne $Global:RDS_LastSelected.NtAccountName) {
        $Job = Get-Job -Name "Mantis_LVServerADProp_$($mantis.SelectedDomain.Name)" -ea 0
        if ((!$Job -or $job.PSBeginTime -lt (Get-Date).AddSeconds(-10))){
            $Job | Remove-Job -Force
            $Global:ControlHandler['ServerNameGbx'].Text = $Global:RDS_LastSelected.ComputerName
            $Global:ControlHandler['LVServerADProp'].Enabled = $false
            $Global:ControlHandler['LVServerADProp'].Items.Clear()
            
            Start-ThreadJob `
                -Name "Mantis_LVServerADProp_$($mantis.SelectedDomain.Name)" `
                -InitializationScript {} `
                -ScriptBlock {
                    param($item)
                    Import-Module ActiveDirectory,PsWrite,RDS.Mantis
                    # $item | Write-Object -PassThru
                    $item | Update-MantisServerProp
                } -ArgumentList $Global:RDS_LastSelected
        }
    }
}
function Update-MantisUserGroups {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $null # Global:ControlHandler['LVServerADProp']
    )
    begin {
        $LstView = @()
        $GroupType = @{
            0x80000002 = 'Security (Global)'
            0x80000004 = 'Security (DomainLocal)'
            0x80000008 = 'Security (Universal)'
            0x00000002 = 'Distribution (Global)'
            0x00000004 = 'Distribution (DomainLocal)'
            0x00000008 = 'Distribution (Universal)'
            0x80000000 = 'Security'
            0x00000000 = 'Distribution'
        }
    }
    process {
        try {
            $Groups = [adsi]"LDAP://$((Get-ADUser ($Items.NtAccountName.split('\')[1]) -Server $Items.domain).DistinguishedName)" | Get-AllMemberOf -Recurse
            $Groups.MemberOf | ForEach-Object {
                [PSCustomObject]@{
                    FirstColValue = $_.properties.name
                    NextValues    = $_.properties.member.count
                    Group         = $GroupType[$_.properties.groupType]
                    Caption       = "Groupe Direct!`nMembres:`n`t$(($_.properties.member | %{($_ -split(','))[0] -replace('CN=')} | Sort-Object -Unique) -join ("`n`t"))"
                    Tag           = $_.properties.distinguishedName

                }
            }
            $Groups.NestedMemberOf | ForEach-Object {
                [PSCustomObject]@{
                    FirstColValue = $_.properties.name
                    NextValues    = $_.properties.member.count
                    Group         = $GroupType[$_.properties.groupType]
                    Caption       = "Groupe InDirect!`nMembres:`n`t$(($_.properties.member | %{($_ -split(','))[0] -replace('CN=')} | Sort-Object -Unique) -join ("`n`t"))"
                    Tag           = $_.properties.distinguishedName
                    shadow        = $True
                }
            }
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }
    }
    end {
        $LstView
    }
}
function Update-MantisServerProp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $null # Global:ControlHandler['LVServerADProp']
    )
    begin {
        $LstView = @()
    }
    process {
        try {
            $ADProperties = Get-ADComputer ($Items.ComputerName -replace(".$($Items.domain)")) -Properties * -Server $Items.Domain
            $ADProperties.PSObject.Properties | ForEach-Object {
                [PSCustomObject]@{
                    FirstColValue = $_.Name
                    NextValues    = ($ADProperties.($_.Name) -join (', '))
                    Group         = "Active Directory"
                    Caption       = ($ADProperties.($_.Name) -join ("`n"))
                }
            }
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }
    }
    end {
        $LstView
    }
}
function Update-MantisUserProp {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $True, ValueFromPipeline = $true)]$Items,
        $Target = $null # Global:ControlHandler['LVUserADProp']
    )
    begin {
        $LstView = @()
        $ADFirst = @('DisplayName','Type','Tel_interne','Phone','Title','Ville','CodePostal')
        $ADEnd = @('Sid','Status','Prenom','Nom','Name','State','Password','Description','PasswordDate','SiteGeo','Service','Bureau','OU','Path','DistinguishedName','CreateDate','Expire','LastLogonDate')
    }
    process {
        try {
            $ADProperties = Get-ADUser ($Items.NtAccountName.split('\')[1]) -Properties * -Server $Items.domain
            $Properties = $ADProperties | Convert-AdUsers
            $Properties.Keys | Where-Object {
                $ADFirst -contains $_
            } | ForEach-Object {
                [PSCustomObject]@{
                    FirstColValue = $_
                    NextValues    = $Properties.$_ -join (', ')
                    Group         = "Active Directory"
                    caption       = $Properties.$_ -join ("`n")
                }
            }
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }

        if ($Items.NtAccountName -match '\w.+\\\w.+') { # -and $Global:ControlHandler['UserNameGbx'].Text -ne $Items.NtAccountName
            try {
                $MailBox = Get-MailInfos $Items.NtAccountName
                if ($MailBox.PrimaryAdress) {
                    $GrpMBox = 'MailBox' # "MailBox $($Mailbox.Alias) on $($Mailbox.DomainName)"
                    $GrpProvider = "Mail provided by $($MailBox.provider.type)"
                    # Start-Process powershell.exe -ArgumentList "-noexit [void](Import-PSSession (Get-RmExchSession $ad) -wa 0); cls; Get-Mailbox -ResultSize 10 -wa 0; Write-Host -back darkgreen 'Vous disposez maintenant de commande Exchange sur $ad!'"
                    [PSCustomObject]@{
                        FirstColValue = 'Email'
                        NextValues    = $Mailbox.PrimaryAdress, $Mailbox.OthersAdresses
                        Group         = $GrpMBox
                        Caption       = "Send and Receive with it"
                    }
                    if ($Mailbox.OthersAdresses) {
                        [PSCustomObject]@{
                            FirstColValue = 'Other Emails'
                            NextValues    = "$($Mailbox.OthersAdresses -join(", "))"
                            Group         = $GrpMBox
                            Caption       = "just Receive with it :`n$($Mailbox.OthersAdresses -join("`n  "))"
                            Shadow        = $true
                        }
                    }
                    [PSCustomObject]@{
                        FirstColValue = 'Domain'
                        NextValues    = $Mailbox.DomainName
                        Tag           = "https://mxtoolbox.com/domain/$($Mailbox.DomainName)/"
                        Group         = $GrpProvider
                        Caption       = "Alias :`n$($Mailbox.Alias)"
                    }
                    [PSCustomObject]@{
                        FirstColValue = 'AdressBook'
                        NextValues    = $Properties.AddressBookMembers | % {(($_ -split (','))[0] -split ('='))[-1]}
                        Tag           = $Properties.AddressBookMembers
                        Group         = $GrpProvider
                        Caption       = ($Properties.AddressBookMembers | % {(($_ -split (','))[0] -split ('='))[-1]}) -join("`n")
                    }
                    

                    [PSCustomObject]@{
                        FirstColValue = "Server"
                        NextValues    = $Mailbox.Provider.DNS
                        Tag           = "https://$($Mailbox.Provider.DNS)/ecp/"
                        Group         = $GrpProvider
                        Caption       = "DblClick pour se connecter en admin"
                    }
                    [PSCustomObject]@{
                        FirstColValue = "Autodiscover"
                        NextValues    = "$($Mailbox.Provider.DNS)"
                        Tag           = "https://mxtoolbox.com/SuperTool.aspx?action=blacklist%3a$($Mailbox.DomainName)&run=toolpage"
                        Group         = $GrpProvider
                        Caption       = "DblClick pour Check Blacklist"
                        Status        = $(if ($Mailbox.Provider.DNS -notmatch '^_autodiscover\._tcp') { 'Warning' })
                    }
                    [PSCustomObject]@{
                        FirstColValue = "MX"
                        NextValues    = "$($Mailbox.Provider.MX -join(", "))"
                        Tag           = "https://mxtoolbox.com/SuperTool.aspx?action=mx%3a$($Mailbox.DomainName)&run=toolpage"
                        Group         = $GrpProvider
                        Caption       = "DblClick pour les details des MX"
                        Status        = $(if ($Mailbox.Provider.MX -notlike 'mx?.coaxis.com') { 'Warning' })
                    }
                    # [PSCustomObject]@{
                    #     FirstColValue    = "Registrar"
                    #     NextValues  = "$($Mailbox.Provider.Registrar -join(", "))"
                    #     Tag = "https://mxtoolbox.com/domain/$($Mailbox.DomainName)/"
                    #     Group   = $GrpProvider
                    #     Caption = "DblClick pour le Test Domain Report"
                    #     # Status  = $(if($Mailbox.Provider.MX -notlike 'mx?.coaxis.com'){'Warning'})
                    # }
                    [PSCustomObject]@{
                        FirstColValue = "Bal [$([math]::round($Mailbox.Bal.Total / 1Gb)) Go]"
                        NextValues    = "$(Convert-ToProgressBarre -block 'l' -length 53 ($Mailbox.Bal.Percent/100))"
                        Group         = $GrpMBox
                        Caption       = "Database : $($Mailbox.Bal.DataBase)`nEspace Alloue : $([math]::round($Mailbox.Bal.Total / 1Gb)) Go`nUtilise [$([math]::round($Mailbox.Bal.Percent))%] : $([math]::round($Mailbox.Bal.Used / 1Gb,2)) Go"
                        Status        = $(if ($Mailbox.Bal.Percent -gt 90) { 'Warning' })
                        Shadow        = $(if (!$Mailbox.Bal.Total) { $true })
                    }
                    [PSCustomObject]@{
                        FirstColValue = "Attachment [$([math]::round($Mailbox.Bal.AttachmentSize / 1Gb)) Go]"
                        NextValues    = "$(Convert-ToProgressBarre -block 'l' -length 53 ($Mailbox.Bal.AttachmentSize/$Mailbox.Bal.Total))"
                        Group         = $GrpMBox
                        Caption       = "Utilise [$([math]::round(($Mailbox.Bal.AttachmentSize/$Mailbox.Bal.Total)*100),1)%] : $([math]::round($Mailbox.Bal.AttachmentSize / 1Gb,2)) Go"
                        Status        = $(if (($Mailbox.Bal.Trash.Size / $Mailbox.Bal.Total) -gt 0.2) { 'Warning' })
                        Shadow        = $(if (!$Mailbox.Bal.Total) { $true })
                    }
                    [PSCustomObject]@{
                        FirstColValue = "Trash [$([math]::round($Mailbox.Bal.Trash.Size / 1Gb)) Go]"
                        NextValues    = "$(Convert-ToProgressBarre -block 'l' -length 53 ($Mailbox.Bal.Trash.Size/$Mailbox.Bal.Total))"
                        Group         = $GrpMBox
                        Caption       = "$($Mailbox.Bal.Trash.Count) elements dans la corbeille"
                        Status        = $(if (($Mailbox.Bal.Trash.Size / $Mailbox.Bal.Total) -gt 0.1) { 'Warning' })
                        Shadow        = $(if (!$Mailbox.Bal.Total) { $true })
                    }
                    if ($Mailbox.Archives) {
                        [PSCustomObject]@{
                            FirstColValue = "Archive [$([math]::round($Mailbox.Archives.Total / 1Gb)) Go]"
                            NextValues    = "$(Convert-ToProgressBarre -block 'l' -length 53 ($Mailbox.Archives.Percent/100))"
                            Group         = $GrpMBox
                            Caption       = "Database : $($Mailbox.Archives.DataBase)`Policy : $($Mailbox.Archives.Policy) jours`nEspace Alloue : $([math]::round($Mailbox.Archives.Total / 1Gb)) Go`nUtilise [$([math]::round($Mailbox.Archives.Percent))%] : $([math]::round($Mailbox.Archives.Used / 1Gb,2)) Go"
                            Status        = $(if ($Mailbox.Archives.Percent -gt 90) { 'Warning' })
                            Shadow        = $(if (!$Mailbox.Archives.Total) { $true })
                        }
                    } else {
                        [PSCustomObject]@{
                            FirstColValue = "Archive [0 Go]"
                            NextValues    = "$(Convert-ToProgressBarre -block 'l' -length 53 0)"
                            Group         = $GrpMBox
                            Caption       = "Aucune strategie d'archives Exchange!"
                            Tag           = 'https://docs.microsoft.com/fr-fr/exchange/security-and-compliance/modify-archive-policies'
                            Shadow        = $true
                        }
                    }
                    if ($Mailbox.Devices) {
                        foreach ($device in $Mailbox.Devices) {
                            [PSCustomObject]@{
                                FirstColValue = $device.name
                                NextValues    = $device.lastSync
                                Group         = "$GrpMBox Devices"
                                Shadow        = $(if ($device.lastSync -lt (get-date).AddDays(-15)) { $true })
                            }
                        }
                    }
                }
            } catch {
                Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber) %Calller%" "", $_ error
            }
        }

        try {
            $Properties.Keys | Where-Object {
                $ADEnd -contains $_
            } | ForEach-Object {
                [PSCustomObject]@{
                    FirstColValue = $_
                    NextValues    = $Properties.$_ -join (', ')
                    Group         = "Other Properties"
                    caption       = $Properties.$_ -join ("`n")
                }
            }
        } catch {
            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
        }
    }
    end {
        $LstView
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
                NextValues = @($_.IP,($_.OperatingSystem -replace('Windows Server ')),$_.InstallDate,$_.RDS.ProductVersion)
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
                # $lastAdded.Tag = $item.DistinguishedName # Tag n'existe pas dans Row
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
                Domain        = $mantis.SelectedDomain.name
                ComputerName  = $($Row.cells[0].value)
                SessionID     = $( try { [int]$Row.cells[1].value } catch { $null } )
                State         = $($Row.cells[2].value)
                Sid           = $(if ($Row.cells[3].value -like 'S-1-*') {$sid = $Row.cells[3].value} else {$sid = $null})
                NtAccountName = $($Row.cells[4].value)
                IPAddress     = $($Row.cells[5].value)
                ClientName    = $($Row.cells[6].value)
                Protocole     = $($Row.cells[7].value)
                handler       = $Row # la ligne dans le DataGridView
            }
        }
    }
    end {
        
    }
}
function Convert-DGV_AD_Row {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]$Rows = $Global:ControlHandler['DataGridView_ADAccounts'].SelectedRows 
    )
    
    begin {}
    process {
        foreach ($Row in $Rows) {
            # Write-Object $Row.cells
            [pscustomobject][ordered]@{
                Domain        = $mantis.SelectedDomain.name
                NtAccountName = $($Row.cells[0].value)
                email         = $($Row.cells[2].value)
                OU            = $($Row.cells[7].value)
                Sid           = $($Row.cells[8].value)
                handler       = $Row # la ligne dans le DataGridView
            }
        }
    }
    end {}
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
function Stop-RDSSessions {
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
function Send-MessageToRDSSessions {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName = $true)] $ComputerName = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $SessionID = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $NtAccountName = $null
    )
    begin {
        $abort = $false
        $Prompt = Input-Box  -Title "Envoyer un popup à l'écran" `
            -Message 'Rédiger un message:' `
            -Value "IMPORTANT: Une maintenance va être appliquée sur les serveurs dans 10 minutes.`nMerci d'enregistrer votre travail et de fermer votre session."
        # -listComboBox @('Info (sans reponse)','AbortRetryIgnore','OK','OKCancel','RetryCancel','YesNo','YesNoCancel')
        if ($Prompt.OUT -eq 'Cancel') {
            $abort = $true
        }
        $CassiaSession = @()
    }
    process {
        if (!$abort) {
            if (!$All) {
                $CassiaSession += $ComputerName | Get-RdComputer | Get-RdSession -Identity $SessionID
            }
            else {
                $Computers += $ComputerName
            }
        }
    }
    end {
        if (!$abort) {
            if ($All) {
                $CassiaSession = Get-RdSession -ComputerName ($Computers | Sort-Object -Unique)
            }
            $footer = "$(whoami), Service Informatique"
            $CassiaSession | Send-RdSession -Message $Prompt.Value -Footer $footer
        }
    }
}
function Request-RDSToSessions {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipelineByPropertyName = $true)] $ComputerName = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $SessionID = $null,
        [Parameter(ValueFromPipelineByPropertyName = $true)] $NtAccountName = $null
    )
    begin {
        $ObjSessions = @()
        $list = @()
        $AskAQuestionToUser = {
            param($Footer, $Question, $ObjSession)
            # import-module PSRdSessions
            # Include Area #
            $SessionAnswer = Request-RdSession -ComputerName $ObjSession.ComputerName -Identity $ObjSession.Identity -Message $Question -Footer $Footer -TimeOut 120
            if ($SessionAnswer.answer -eq 'no') {
                $icon = [System.Windows.Forms.ToolTipIcon]::Error
                # $ObjSessions.handler | Write-Object -fore red
            } else {
                $icon = [System.Windows.Forms.ToolTipIcon]::Info
                # $ObjSessions.handler | Write-Object -fore red
            }
            # pop up windows en bas à droite de l'écran
            # s'assurer que l'assistant de concentration n'est pas activé sinon la notification n'apparait pas comme escomptée
            Invoke-BalloonTip -Timeout 300 -Title "Réponse de [$($SessionAnswer.UserAccount)]" -type $icon -Message "$($SessionAnswer.Request)`n[ $($SessionAnswer.answer) ]"
            [PSCustomObject]@{
                Request         = $SessionAnswer.Request
                ComputerName    = $SessionAnswer.Server.ServerName
                UserAccount     = $SessionAnswer.UserAccount
                Answer          = $SessionAnswer.Answer
                ClientName      = $SessionAnswer.ClientName
                ClientIPAddress = $SessionAnswer.ClientIPAddress
            }
        } | Include-ToScriptblock -Functions (Get-Command -Name 'PsBright\Invoke-BalloonTip', 'PSRdSessions\Request-RdSession')
    }
    process {
        # Write-Object $Session -foreGroundColor Red
        $ObjSessions += [PSCustomObject]@{
            ComputerName  = $ComputerName
            Identity      = $SessionID
            NtAccountName = $NtAccountName
        }
    }
    end {
        $list = @($ObjSessions | ForEach-Object { $_.NtAccountName }) -join ("`n")

        $input_box = Input-Box -Title "Envoyer une question dichotomique (oui/non) à l'écran" `
            -Message "Poser une question à ces utilisateurs : `n$list" `
            -Value "Pouvez-vous svp me confirmer que ce problème est résolu ?"

        if ($input_box.OUT -ne 'Cancel') {
            $question = $input_box.value
            $footer = "$(whoami), Service Informatique"
            $ObjSessions | ForEach-Object {
                # c'est cela qui démarre l'action en asynchrone
                Start-ThreadJob -name "Mantis_AskAQuestionToUser_$($_.Identity)" -ScriptBlock $AskAQuestionToUser -ArgumentList @($footer, $question, $_)
            }
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