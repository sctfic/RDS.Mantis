
# $srv = [ServersCollector]::new()
# $Srv.Get()
class UsersCollector {
    hidden $Domain = $null
    hidden $Job = $null
    $Items = $null

    UsersCollector ($Domain){
        $this.Domain = $Domain
        $this.Get()
    }
    UsersCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
        $this.Get()
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            if(!$this.Job){
                $this.Job = Start-ThreadJob `
                    -Name "Srv_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Get-ADUser -Identity 'alopez' `
                            -Properties proxyaddresses,msexchhomeservername,msexchhomeservername,homemdb,msexchdumpsterquota,msexchdelegatelistbl,msexcharchivename,msexcharchivedatabaselink,msexchmailboxtemplatelink,msexcharchivequota `
                            -Filter {enabled -eq $true} `
                            -Server $Domain | %{
                                try {
                                } catch {
                                    Write-Error $_
                                    # Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", "Impossible de determiner l'adresse IP de [$DNSHostName]" error
                                }
                                [PSCustomObject]@{
                                    Name = $_.DNSHostName
                                    DN = $_.DistinguishedName
                                    SID = $_.SID
                                    OperatingSystem = $_.OperatingSystem
                                    IP = $IP
                                }
                            }
                    } `
                    -ArgumentList $this.Domain
            } else {
                $this.Items = $this.Job | Receive-Job -AutoRemoveJob -Wait
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
        $this.Get() # first call when creating Class object
    }
    ServersCollector (){
        $this.Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain()).Forest.Name
        $this.Get() # first call when creating Class object
    }
    [PSCustomObject[]] Get () {
        if (!$this.Items) {
            if(!$this.Job){ # on first call, just start thread
                $this.Job = Start-ThreadJob `
                    -Name "Srv_$($this.Domain)" `
                    -InitializationScript {} `
                    -ScriptBlock {
                        param($Domain)
                        Get-ADComputer -Filter { OperatingSystem -Like '*Windows Server*'} -Properties OperatingSystem,whencreated -Server $Domain | %{
                            $DNSHostName = $_.DNSHostName
                            $IP = $null
                            try {
                                $IP = [string][System.Net.Dns]::GetHostAddresses($DNSHostName).IPAddressToString -Split(' ') | Sort-Object -Unique
                            } catch {
                                Write-Error $_
                                Write-Error "Impossible de determiner l'adresse IP de [$DNSHostName]"
                                # Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", "Impossible de determiner l'adresse IP de [$DNSHostName]" error
                            }
                            [PSCustomObject]@{
                                Name = $_.DNSHostName
                                DN = $_.DistinguishedName
                                SID = $_.SID
                                OperatingSystem = $_.OperatingSystem
                                IP = $IP
                            }
                        }
                    } `
                    -ArgumentList $this.Domain
            } else { # on secand call, wait and read started thread
                $this.Items = $this.Job | Receive-Job -AutoRemoveJob -Wait
                # Write-Host $this.Items -ForegroundColor DarkYellow
            }
        }
        return $this.Items
    }
}

class Mantis {
    $CurrentDomain = $null
    $TrustedDomain = $null

    Mantis () {
        $this.GetDomains() | Out-Null
        # $this.Get() | Out-Null
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
                Users = [UsersCollector]::new($Current)
            }
            $Trusted += ([System.DirectoryServices.ActiveDirectory.Forest]::GetCurrentForest().GetAllTrustRelationships()).targetName
            $Trusted += ([System.DirectoryServices.ActiveDirectory.Domain]::getCurrentDomain().GetAllTrustRelationships()).targetName
            $this.TrustedDomain = $Trusted | ?{$_ -and (!$reachable -or (Test-TcpPort $_ -port 135 -timeout 200 -ConfirmIfDown -Quick))} | %{
                [PSCustomObject]@{
                    DistinguishedName = "DC=$($_ -replace('\.',',DC='))"
                    ShortName = ($_ -split('\.'))[0]
                    Name = $_
                    Servers = [ServersCollector]::new($_)
                    Users = [UsersCollector]::new($_)
                }
            }
        }
        return @($this.CurrentDomain)+$this.TrustedDomain
    }
    [PSCustomObject[]] Domain ($Name) {
        return $this.GetDomains() | Where-Object{
            $_.Name -like $Name
        }
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