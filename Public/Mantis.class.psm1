
# $srv = [ServersCollector]::new()
# $Srv.Get()
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
                $this.Job = Start-ThreadJob -StreamingHost $Global:host `
                    -Name "Mantis_Usr_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector                         Get-ADUser',$Domain -mode wait
                        Get-ADUser `
                            -Properties proxyaddresses,msexchhomeservername,msexchhomeservername,homemdb,msexchdumpsterquota,msexchdelegatelistbl,msexcharchivename,msexcharchivedatabaselink,msexchmailboxtemplatelink,msexcharchivequota `
                            -Filter {enabled -eq $true} `
                            -Server $Domain
                    } -ArgumentList $this.Domain
            } else {
                $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
                # Write-Host $this.Items -ForegroundColor DarkYellow
            }
        }
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
                $this.Job = Start-ThreadJob -StreamingHost $Global:host `
                    -Name "Mantis_Grp_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector                         Get-ADGroup',$Domain -mode wait
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
                $this.Job = Start-ThreadJob -StreamingHost $Global:host `
                    -Name "Mantis_Srv_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Write-LogStep 'Collector Get-ADComputer',$Domain -mode wait
                        Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*'} -Properties OperatingSystem,whencreated -Server $Domain | ForEach-Object -Parallel {
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
                            if ($IP -and (Get-Registry "\\$($_.DNSHostName)\HKLM\SYSTEM\CurrentControlSet\Control\Terminal Server").Type -eq 'Container') {
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
                                OperatingSystem = $_.OperatingSystem
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
                            }
                        } -ThrottleLimit 12
                    } `
                    -ArgumentList $this.Domain
            } else { # on secand call, wait and read started thread
                $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
                # Write-Host $this.Items -ForegroundColor DarkYellow
            }
        }
        return $this.Items
    }
}
class DFSCollector {
    hidden $Domain = $null
    # hidden $Job = $null
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
            #     $this.Job = Start-ThreadJob -StreamingHost $Global:host `
            #         -Name "Mantis_DFS_$($this.Domain)" `
            #         -InitializationScript {} `
            #         -ScriptBlock {
            #             param($Domain,$Root)
            #             Get-DfsnRoot $Domain # tres lent ! | ForEach-Object {
            #                 [PSCustomObject]@{
            #                     Name = Split-Path -Leaf $_
            #                     Path = $_
            #                 }
            #             }
            #         } -ArgumentList $this.Domain,$this.Root
            # } else {
            #     $this.Items = $this.Job | Receive-Job -Wait -AutoRemoveJob
            #     # Write-Host $this.Items -ForegroundColor DarkYellow
            # }
        }
        return $this.Items
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


function Get-Mantis {
    [Mantis]::new()
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