
# $srv = [ServersCollector]::new()
# $Srv.Get()
class Sequence {
    hidden $Cursor = 0
    hidden $ListScriptBlock = $null
    Sequence ($ListScriptBlock) {
        $this.ListScriptBlock = $ListScriptBlock
    }
    [void]Append ($ScriptBlock){
        $this.ListScriptBlock += $ScriptBlock
    }
    [scriptblock]Next(){
        $Next = $Null
        if ($this.Cursor -lt $this.ListScriptBlock.count){
            Write-logstep "Sequence [$($this.Cursor)]",$this.ListScriptBlock[$this.Cursor] OK
            $this.Cursor++
            $Next = $this.ListScriptBlock[$this.Cursor-1]
        }
        return $Next
    }
    [void]Restart(){
        $this.Cursor = 0
        # $this.ListScriptBlock = $null
    }
}
class UsersCollector {
    hidden $Domain = $null
    hidden $Job = $null
    $Items = $null

    UsersCollector ($Domain){
        $this.Domain = $Domain
    }
    UsersCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            if(!$this.Job){
                # $this.Job = $true
                #    -StreamingHost $Global:host `
                $this.Job = Start-ThreadJob `
                    -Name "Mantis_Usr_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector Get-ADUser',$Domain -mode wait
                        Get-ADUser `
                            -Properties * `
                            -Filter {enabled -eq $true} `
                            -Server $Domain | Convert-AdUser
                    } -ArgumentList $this.Domain
            } else {
                $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
                # Write-Host $this.Items -ForegroundColor DarkYellow
            }
        }
        # if($This.Job.status -eq 'Completed') {Remove-Job $this.Job}
        return $this.Items
    }
}
class GroupsCollector {
    hidden $Domain = $null
    hidden $Job = $null
    $Items = $null

    GroupsCollector ($Domain){
        $this.Domain = $Domain
    }
    GroupsCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            if(!$this.Job){
                # $this.Job = $true
                #    -StreamingHost $Global:host `
                $this.Job = Start-ThreadJob `
                    -Name "Mantis_Grp_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector Get-ADGroup',$Domain -mode wait
                        Get-ADGroup `
                            -Properties Description,member,MemberOf,Members `
                            -Filter '*' `
                            -Server $Domain
                    } -ArgumentList $this.Domain
            } else {
                $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
                # Write-Host $this.Items -ForegroundColor DarkYellow
            }
        }
        # if($This.Job.status -eq 'Completed') {Remove-Job $this.Job}
        return $this.Items
    }
}
class DFSCollector {
    hidden $Domain = $null
    hidden $Job = $null
    $Root = $null
    $Items = $null

    DFSCollector ($Domain){
        $this.Domain = $Domain
        $this.Root = "\\$($this.Domain)"
    }
    DFSCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
        $this.Root = "\\$($this.Domain)"
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            $this.Items = (Get-ADObject -Server $this.Domain -SearchBase "CN=Dfs-Configuration,CN=System,DC=$($this.Domain -replace('\.',',DC='))" -Filter * -SearchScope OneLevel).Name | ForEach-Object {
                [PSCustomObject]@{
                    Name = $_
                    Path = "$($this.Root)\$_"
                }
            }
            # if(!$this.Job){
            #     $this.Job = $true
            #    -StreamingHost $Global:host `
            #     $this.Job = Start-ThreadJob `
            #         -Name "Mantis_DFS_$($this.Domain)" `
            #         -InitializationScript {} `
            #         -ScriptBlock {
            #             param($Domain,$Root)
            #             Get-DfsnRoot $Domain | ForEach-Object { # tres lent ! et semble avoir des PB en mode thread
            #                 [PSCustomObject]@{
            #                     Name = Split-Path -Leaf $_.Path
            #                     Path = $_.Path
            #                 }
            #             }
            #         } -ArgumentList $this.Domain,$this.Root
            # } else {
            #     Write-Host 'Wait !!!' -fore DarkYellow
            #     $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
            #     Write-Host 'Wait ok !' -fore DarkGreen
            # }
        }
        # if($This.Job.status -eq 'Completed') {Remove-Job $this.Job}
        return $this.Items
    }
}

class ServersCollector {
    hidden $Domain = $null
    hidden $Job = $null
    $Items = $null

    ServersCollector ($Domain){
        $this.Domain = $Domain
    }
    ServersCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            if(!$this.Job){ # on first call, just start thread
                $this.Job = Start-ThreadJob `
                    -StreamingHost $Global:host `
                    -Name "Mantis_Srv_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector Get-ADComputer',$Domain -mode wait
                        Get-ADComputer -Filter { operatingSystem -Like '*Windows Server*'} -Properties whencreated -Server $Domain | ForEach-Object -Parallel {
                            function Write-LogStep { }
                            $DNSHostName = $_.DNSHostName
                            $IP = $null
                            try {
                                $IP = [string][System.Net.Dns]::GetHostAddresses($DNSHostName).IPAddressToString -Split(' ') | Sort-Object -Unique
                            } catch {
                                Write-Error $_
                                Write-Error "Impossible de determiner l'adresse IP de [$DNSHostName]"
                                # Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", "Impossible de determiner l'adresse IP de [$DNSHostName]" error
                            }
                            if ($IP -and (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server").Type -eq 'Container' -and (Test-TcpPort $_.DNSHostName -Quick -ConfirmIfDown)) {
                                $OperatingSystem = (Get-Registry "\\$($_.DNSHostName)\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\ProductName").value
                                $CurrentBuild = (Get-Registry "\\$($_.DNSHostName)\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\CurrentBuild").value
                                $ProductVersion = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\ProductVersion").value
                                $fDenyTSConnections = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\fDenyTSConnections").value
                                $TSUserEnabled = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\TSUserEnabled").value
                                $TSEnabled = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server\TSEnabled").value
                                # if ($TSEnabled) {
                                    $MemberOfFarm = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings\SessionDirectoryClusterName").value
                                    $ServerBroker = (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\ControlSet001\Control\Terminal Server\ClusterSettings\SessionDirectoryLocation").value
                                # }
                            }

                            [PSCustomObject]@{
                                Name = $_.DNSHostName
                                DN = $_.DistinguishedName
                                SID = $_.SID.value
                                OperatingSystem = $(if($OperatingSystem){"$OperatingSystem [$CurrentBuild]"})
                                InstallDate = $_.whencreated
                                IP = $IP
                                isDC = $($_.DistinguishedName -like '*,OU=Domain Controllers,DC=*')
                                isBroker = (Get-Registry "\\$($_.DNSHostName)\HKLM\SOFTWARE\Microsoft\Windows NT\CurrentVersion\Terminal Server\CentralPublishedResources\PublishedFarms\").Value
                                RDS = [PSCustomObject]@{
                                    ProductVersion = $ProductVersion
                                    fDenyTSConnections = $fDenyTSConnections
                                    TSEnabled = $TSEnabled
                                    TSUserEnabled = $TSUserEnabled
                                    MemberOfFarm = $MemberOfFarm
                                    ServerBroker = $ServerBroker
                                }
                                Sessions = $null # $_.DNSHostName | Get-RdSession | Convert-RdSession
                            }
                        } -ThrottleLimit 12
                    } `
                    -ArgumentList $this.Domain
            } else { # on se$true
                # on next call, wait and read started thread
                $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
                $this.Job = $null
                $this.GetRDSessions()
            }
        }
        # if($This.Job.status -eq 'Completed') {Remove-Job $this.Job}
        return $this.Items
    }
    [PSCustomObject[]] GetRDSessions () {
        return $this.GetRDSessions($this.Items)
    }
    [PSCustomObject[]] GetRDSessions ($ComputerItems) {
        if($ComputerItems){ # on first call, just start thread
            if (!$this.Job) {
                $this.Items | ForEach-Object {
                    $_.Sessions = $null
                }
                $this.Job = Start-ThreadJob `
                    -Name "Mantis_Sessions_$($this.Domain)" `
                    -InitializationScript {Import-Module PSWrite,PSRdSessions -DisableNameChecking -SkipEditionCheck} `
                    -ScriptBlock {
                        param($Computers)
                        $Computers | ForEach-Object -Parallel {
                            $_.Sessions = $_.Name | Get-RdSession | Convert-RdSession
                        } -ThrottleLimit 12 -TimeoutSeconds 20
                    } -ArgumentList @(,$ComputerItems) `
                    # -StreamingHost $Global:host
            } else { # on se$true
                # on next call, wait and wait and read started thread
                $this.Job | Receive-Job -Wait -AutoRemoveJob
                # $this.Job | Write-Object -PassThru
                $this.Job = $null
                return $this.Items.Sessions
            }
        }
        return $null
    }
}


class Mantis {
    $CurrentDomain = $null
    $SelectedDomain = $null
    $TrustedDomain = $null

    Mantis () {
        $this.GetDomains() | Out-Null
        # $this.CurrentDomain.Servers.Get()
        # $this.CurrentDomain.Users.Get()
        # $this.CurrentDomain.Groups.Get()
        # $this.CurrentDomain.Dfs.Get()
        $this.SelectedDomain = $this.CurrentDomain
    }

    [PSCustomObject[]] GetDomains () {
        if (!$this.CurrentDomain){
            $Trusted = @()
            $Current = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
            # Write-Host $Current -ForegroundColor DarkYellow
            $this.CurrentDomain = [PSCustomObject]@{
                DistinguishedName = "DC=$($Current -replace('\.',',DC='))"
                ShortName = ($Current -split('\.'))[0]
                Name = $Current
                Servers = [ServersCollector]::new($Current)
                Groups = [GroupsCollector]::new($Current)
                Users = [UsersCollector]::new($Current)
                DFS = [DFSCollector]::new($Current)
            }
            $Trusted += ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GetAllTrustRelationships()).targetName
            $Trusted += ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain().GetAllTrustRelationships()).targetName
            $this.TrustedDomain = $Trusted | Where-Object{$_ -and (!$reachable -or (Test-TcpPort $_ -port 135 -timeout 200 -ConfirmIfDown -Quick))} | ForEach-Object{
                [PSCustomObject]@{
                    DistinguishedName = "DC=$($_ -replace('\.',',DC='))"
                    ShortName = ($_ -split('\.'))[0]
                    Name = $_
                    Servers = [ServersCollector]::new($_)
                    Groups = [GroupsCollector]::new($_)
                    Users = [UsersCollector]::new($_)
                    DFS = [DFSCollector]::new($_)
                }
            }
        }
        return @($this.CurrentDomain)+$this.TrustedDomain
    }
    [PSCustomObject[]] Domain ($Name) {
        $this.SelectedDomain = $this.GetDomains() | Where-Object{
            $_.Name -like $Name
        }
        # $this.SelectedDomain.Servers.Get()
        # $this.SelectedDomain.Users.Get()
        # $this.SelectedDomain.Groups.Get()
        # $this.SelectedDomain.Dfs.Get()
        return $this.SelectedDomain
    }
}


function New-Mantis {
    [Mantis]::new()
}
function New-Sequence($s) {
    [Sequence]::new($s)
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