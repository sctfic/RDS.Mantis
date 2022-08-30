$PSModuleAutoloadingPreference = 1
function Convert-AdUsers {
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
                $AdsiItem = [adsi]"LDAP://$($item.DistinguishedName)"
                # write-host -ForegroundColor Green (Measure-Command {
                    $rds = $null
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
                        $TsAllowLogon = 'Limited Account!'
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
                    Status               = $(if($Item.Enabled){''}else{'Disabled'})
                    Phone                = "$($Item.telephonenumber)"
                    Prenom               = "$($Item.givenname)"
                    Nom                  = "$($Item.sn)"
                    Name                 = "$($Item.Name)"
                    State                = "$($Item.State)"
                    Tel_interne          = "$($Item.ipPhone)"
                    Password             = $null
                    Title                = "$($item.Title)"
                    AddressBookMembers   = $item.showInAddressBook
                    Description          = "$($Item.description)"
                    # PasswordType         = $PassWordType
                    PasswordDate         = $(try {$item.PasswordLastSet.ToString()} catch {''})
                    Groupes              = $($Item.memberOf | ?{$_} | %{(($_ -split(','))[0] -split('='))[1]}) # -join ("`n") # | ForEach-Object {([adsi]"LDAP://$_").name}
                    CodePostal           = "$($Item.postalcode)"
                    Ville                = "$($Item.City)"
                    SiteGeo              = "$($Item.State)"
                    Service              = "$($Item.Department)"
                    Bureau               = "$($Item.physicaldeliveryofficename)"
                    OU                   = "$($Itemadsi.Path -replace('LDAP://'))"
                    Path                 = "LDAP://$($Item.DistinguishedName)"
                    DistinguishedName    = "$($Item.DistinguishedName)"
                    CreateDate           = $(try {$item.whenCreated.ToString()} catch {''})
                    Expire               = $(try {$item.AccountExpirationDate.ToString()} catch {''})
                    LastLogonDate            = $(try {$item.LastLogonDate.ToString()} catch {''})
                }
                $Prop
            } catch {
                Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber) %Caller%" '',$_ error
            }
        }
    }
    end {
    }
}
function Get-AllMemberOf {
    Param(
        [Parameter(ValueFromPipeline = $true)]$ADSIObject,
        [switch]$Recurse,
        $Groups = $null
    )
    begin {
    }
    process {
        if ($null -eq $Groups) {
            $Groups = [PSCustomObject]@{
                MemberOf       = @()
                NestedMemberOf = @()
            }
            Foreach ($Memberof in $ADSIObject.properties.memberof) {
                $Group = [adsi]"LDAP://$MemberOf"
                $Groups.Memberof += $group
                if ($Group.properties.memberof -and $Recurse) {
                    $Groups = $group | Get-ADSIMemberOf -Groups $Groups -Recurse
                }
            }
        } else {
            Foreach ($Memberof in $ADSIObject.properties.memberof) {
                $path = "LDAP://$MemberOf"
                if ($Groups.MemberOf.path -contains $path) {
                } elseif ($Groups.NestedMemberOf.path -contains $path) {
                } else {
                    $Group = [adsi]$path
                    $Groups.NestedMemberof += $group
                    $Groups = $group | Get-ADSIMemberOf -Groups $Groups -Recurse
                }
            }
        }
        $Groups
    }
    end {
    }
}
# PSWinForm-Builder\New-WinForm -DefinitionFile "$PSScriptRoot\GUI\MantisForm.psd1" -Verbose -PreloadModules PsWrite,rds.mantis

