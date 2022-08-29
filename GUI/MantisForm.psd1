    @{  ControlType = 'Form'
    # Name            = 'MainForm'
    # ipmo -Force PSWinForm-Builder
    # New-WinForm -DefinitionFile C:\Users\alopez\Documents\PowerShell\Modules\RDS.Mantis\MantisForm.psd1 -PreloadModules PsWrite,PsBright -Verbose | Out-Null
    Size            = '1280, 800'
    Text            = 'RDS.Mantis'
    # TopMost         = $true
    Anchor          = 'Right,Top'
    # AutoSize        = $False
    MaximizeBox     = $true
    MinimizeBox     = $False
    ControlBox      = $True # True or False to show close icon on top right corner
    # FormBorderStyle = 'FixedSingle' # Fixed3D, FixedDialog, FixedSingle, Sizable, None, FixedToolWindow, SizableToolWindow
    KeyPreview      = $true
    Icon            = [System.Drawing.Icon][System.IO.MemoryStream] [System.Convert]::FromBase64String('AAABAAEALUAQAAEABABoCAAAFgAAACgAAAAtAAAAgAAAAAEABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAWFxsAHh8jACkqLQA4ODwARkZJAFRVWABkZGcAc3R3AIGChQCRkpUAo6WnALW2uQDGyMsA4eLlAPj5+gAAAAAA//////////////////////////////AA////////////////////////f/////AA////////////////////////+f////AA////93//////////////////h/////AA/////0b6///////////////79f////AA/////2KPb/////////////+hEa////AA//////Rf9//////////////2QP////AA//////gGxI/P//////////90MY////AA//////8i+D/5///////////4IU////AA//////9xfwf4f/////////9nYS////AA///////wOQL/b//////////yEQn///AA//////9xFhF/dt////////+rIR////AA///////xERE/8v////////o0IQX///AA//////9REREPkD////////+RERr///AA//////lBEREV8Qn///////b1ERn///AA///////GERESQRO//////9cBERn//6AA///////2ARERERH////////xERn/9/AA///////2ERERERFP/////2VREQ//c4AA///////xEREREREJ//////EREQj/8XAA/////4UBEhEREREX////j8YRER//MUAA//////EREREREREU////9PcxEU/3gjAA////+RERESIRERERf///9RAREX//gSAA////8hERE/YBERER///3/0EREP/3USAA////UREQf/9BERIRf//4BWERE7+5UTAA///2ESEW///yERERL/r/QRERFf/1IUAA//+SERGP///4ERERD/hfkRERD/9ZEXAA//9RERr/////YRERF/sBYRERb/+BEYAA//8REJ//////ghEREH9BERERn7+2EIAA//URCf//////9hEREReBEREG//gCEvAA/5MRb////////0EREREhERFP/2lRFPAA//AW//////////IREREREhL/+fERB/AA/3E///////////gBESEREQj/xoURP/AA/2H///////////9REREREf/09wER//AA/6e///+fYj////9xERERBv9RFREUn/AA//////YwERj////xERER//YBEREY//AA////dRERERKP//9yAREm/8MRERFv//AA///2ERERERFv///zEhP///YREQT///AA//8jZgERERFf//+SEl////8BEl////AA//Kt7oERERfo///xSf////8hJ/////AA+F3u7qERERvpz/+vj/////82r/////AA9N7u7UERERrr//////////////////AAjO7u5xERERjsn/////////////////AAfu7ugBERERPb//////////////////AAfe7XERERERGXr/////////////////AA9ZlRERERAAAp//////////////////AA/5h2Ylkzn6/P//////////////////AA////9G9y//////////////////////AA////g69hf/////////////////////AA////8v/y//////////////////////AA////Ym/Eb/////////////////////AA////8//5L/////////////////////AA////k5//Kf////////////////////AA////9f//9P////////////////////AA////1H//84////////////////////AA////9P//+U////////////////////AA////95///2f///////////////////AA////9/////X///////////////////AA/////1////9f//////////////////AA/////3/////2//////////////////AA/////2//////b/////////////////AA////////////+P////////////////AA/////////////5/v//////////////AA//////r///////////////////////AA///////P//////////////////////AA///////4AAD/////9/gAAP/////7+AAA/n////P4AAD/L///6/gAAP8X///B+AAA/5v//+P4AAD/gX//wfgAAP/Jv//h+AAA/8Sf/8H4AAD/4N//4PgAAP/AR//B+AAA/+Bv/4D4AAD/wEf/wPgAAP+AI/+g+AAA/8AB/wDwAAD/4AP/4OgAAP/gAf8BwAAA/+AA/4DgAAD/AAD9AcAAAP+AAP6BgAAA/gAAfgHAAAD+AgD7A4AAAPwHAHgBAAAA+A+AbAOAAADwH4BkBwAAAPA/wCAHAAAA8H/AEAUAAADg/+AADggAAMH/8AAcCAAA4//4ADoIAADH//gAMBgAAM///ADoOAAAx9H8AMAYAAD/gP4DgDgAAPwAfAMAeAAA+AB+D4D4AADwAHwfwfgAAOAAPj/D+AAAgAAdf8f4AACAAD////gAAAAAH///+AAAAAA////4AAAAAB////gAAIAAP///+AAAwAV////4AAD+T/////gAAPxH////+AAA/u/////4AAD8R/////gAAP7n////+AAA/HP////4AAD++/////gAAPx5////+AAA/vn////4AAD+fP////gAAP7+////+AAA/39////4AAD/f7////gAAP9/3///+AAA///v///4AAD///X///gAAP+/////+AAA/9/////4AAA=')
    Events      = @{
        Load    = [Scriptblock]{ # Event
            $this.TopMost = $true
            Invoke-EventTracer $this 'Load'
            if (-not ('Console.Window' -as [type])) {
                Add-Type -Name Window -Namespace Console -MemberDefinition '[DllImport("Kernel32.dll")]public static extern IntPtr GetConsoleWindow();[DllImport("user32.dll")]public static extern bool ShowWindow(IntPtr hWnd, Int32 nCmdShow);'
            }
            $Script:consolePtr = [Console.Window]::GetConsoleWindow()
            if (!$PSCmdlet.MyInvocation.BoundParameters['Verbose'].IsPresent) {
                [Console.Window]::ShowWindow($Script:consolePtr,0) | Out-Null
            }
            Write-LogStep 'Load',$this.name  ok -logTrace $Global:ControlHandler['RichTextBox_Logs']
            $this.TopMost = $false
        }
        Shown = [Scriptblock]{ # Event
            Invoke-EventTracer $this 'Shown'
            Invoke-MantisConstructor
        }
        KeyDown = [Scriptblock]{ # Event
            if ($_.KeyCode -eq 'Escape') {
                $this.Close()
            } else {
                $Global:KeyDown = $_
            }
        }
        KeyUp = [Scriptblock]{ # Event
            $Global:KeyDown = $null
        }
        Closing = [Scriptblock]{ # Event
            Invoke-EventTracer $this 'Closing'
            if ($Script:consolePtr) {
                [Console.Window]::ShowWindow($Script:consolePtr,5) | Out-Null
            }
            Write-LogStep 'Load',$this.name  ok -logTrace $true
        }
    }
    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
        @{  ControlType = 'TabControl'
            Dock        = 'Fill'
            Events      = @{
                Enter = [Scriptblock]{ # Event
                    Invoke-EventTracer $this 'SelectedIndexChanged'
                    # $Global:ControlHandler['PanelLeft'].Width = 220
                    # $Global:ControlHandler['PanelRight'].Width = 100
                    # $Global:ControlHandler['Logs'].Height = 58
                }
                SelectedIndexChanged = [Scriptblock]{ # Event
                    Invoke-EventTracer $this 'SelectedIndexChanged'
                    switch ($this.SelectedTab.Name) {
                        'Tab1' {
                        }
                        'Tab2' {
                        }
                        default {}
                    }
                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'TabPage'
                    Name        = 'RDSFarm'
                    Text        = 'Sessions RDS'
                    Dock        = 'Fill'
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'DataGridView'
                            Name             = 'DataGridView_Sessions'
                            Dock             = 'Fill'
                            ReadOnly         = $True
                            CellBorderStyle  = 'SingleHorizontal'
                            SelectionMode    = 'FullRowSelect'
                            RowHeadersVisible = $False
                            AllowUserToAddRows = $False
                            AllowUserToResizeRows = $false
                            AllowUserToDeleteRows = $False
                            RowTemplate      = [System.Windows.Forms.DataGridViewRow]@{Height = 20}
                            AlternatingRowsDefaultCellStyle = [System.Windows.Forms.DataGridViewCellStyle]@{BackColor='AliceBlue'}
                            Events      = @{
                                Enter    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Enter'
                                }
                                Click    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Click' # left and right but not on empty area
                                    $global:RDS_LastSelected = $this.SelectedRows[-1] | Convert-DGV_RDS_Row
                                    Start-Job4RightProperties
                                }
                                DoubleClick    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'DoubleClick'
                                    $this.SelectedRows | Convert-DGV_RDS_Row | Start-ActionByRow -ColumnClicked $this.CurrentCell.OwningColumn.HeaderText
                                }
                            }
                            Childrens   = @(
                                @{  ControlType = 'ContextMenuStrip'
                                        Dock        = 'Fill'
                                        Events      = @{
                                            Opening    = [Scriptblock]{ # Event
                                                Invoke-EventTracer $this 'MenuEnter'
                                                $Sessions = $Srv = 0
                                                $global:RDS_Selected = $Global:ControlHandler['DataGridView_Sessions'].SelectedRows | Convert-DGV_RDS_Row
                                                $Srv = ($global:RDS_Selected.ComputerName | Where-Object {
                                                    $_
                                                } | Sort-Object -Unique).count

                                                $Sessions = ($global:RDS_Selected.NtAccountName | Where-Object {
                                                    $_
                                                }).count
                                                $Global:ControlHandler['ContextMenuStrip_RDSessions_Title'].Text = "$Sessions Sessions sur $Srv Serveurs"
                                                Write-Host $Sessions, $Global:ControlHandler['ContextMenuStrip_RDSessions_Title'].Text -ForegroundColor Green
                                            }
                                        }
                                        Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{
                                            ControlType = 'ToolStripLabel'
                                            Name        = 'ContextMenuStrip_RDSessions_Title'
                                            Text        = 'X Sessions sur Y Serveurs'
                                        },
                                        @{ ControlType = 'ToolStripSeparator'
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'Delete'
                                            ShortcutKeyDisplayString = 'Supp.'
                                            Text        = 'Fermer la session'
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $global:RDS_Selected = $Global:ControlHandler['DataGridView_Sessions'].SelectedRows | Convert-DGV_RDS_Row
                                                    if ($global:RDS_Selected | Stop-DGV_RDSSessions) {
                                                        $Global:ControlHandler['DataGridView_Sessions'].SelectedRows | ForEach-Object {
                                                            $Global:ControlHandler['DataGridView_Sessions'].Rows.Remove($_)
                                                        }
                                                        $Global:Mantis.SelectedDomain.Servers.RefreshRDSessions()
                                                    }
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        },
                                        @{ ControlType = 'ToolStripSeparator'
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'Ctrl+M'
                                            ShortcutKeyDisplayString = 'Ctrl+M'
                                            Text        = "Envoyer un Message a l'ecran"
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $global:RDS_Selected = $Global:ControlHandler['DataGridView_Sessions'].SelectedRows | Convert-DGV_RDS_Row
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'Ctrl+Q'
                                            ShortcutKeyDisplayString = 'Ctrl+Q'
                                            Text        = "Envoyer une Question Y/N a l'ecran"
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $global:RDS_Selected = $Global:ControlHandler['DataGridView_Sessions'].SelectedRows | Convert-DGV_RDS_Row
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'F5'
                                            ShortcutKeyDisplayString = 'F5'
                                            Text        = "Actualiser"
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $Global:Mantis.SelectedDomain.Servers.RefreshRDSessions()
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        }
                                    )
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ComputerName'
                                    Width = 140
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'SessionID'
                                    Width = 35
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'State'
                                    Width = 80
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Sid'
                                    Width = 45
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'UserAccount'
                                    Width = 150
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'IPAddress'
                                    Width = 125
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ClientName'
                                    Width = 130
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Protocole'
                                    Width = 75
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ClientBuildNumber'
                                    Width = 60
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'LoginTime'
                                    Width = 115
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ConnectTime'
                                    Width = 115
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'DisconnectTime'
                                    Width = 115
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Inactivite'
                                    Width = 100
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Screen'
                                    Width = 100
                                    
                                },
                                @{ ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Process'
                                    Width = 0
                                }
                            )
                        },
                        @{  ControlType = 'ProgressBar'
                            Name        = 'ProgressBar_DataGridView_Sessions'
                            Style       = 'Marquee'
                            Dock        = 'Top'
                            Height      = 5
                        }
                    )
                },
                @{  ControlType = 'TabPage'
                    Name        = 'ADAccounts'
                    Text        = 'Comptes AD'
                    Dock        = 'Fill'
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'DataGridView'
                            Name             = 'DataGridView_ADAccounts'
                            Dock             = 'Fill'
                            ReadOnly         = $True
                            CellBorderStyle  = 'SingleHorizontal'
                            SelectionMode    = 'FullRowSelect'
                            RowHeadersVisible = $False
                            AllowUserToAddRows = $False
                            AllowUserToDeleteRows = $False
                            RowTemplate      = [System.Windows.Forms.DataGridViewRow]@{Height = 20}
                            AlternatingRowsDefaultCellStyle = [System.Windows.Forms.DataGridViewCellStyle]@{BackColor='AliceBlue'}
                            Events      = @{
                                Enter    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Enter'
                                }
                                Click    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Click' # left and right but not on empty area
                                    $global:RDS_LastSelected = $this.SelectedRows[-1] | Convert-DGV_AD_Row
                                    Start-Job4RightProperties
                                }
                                DoubleClick    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'DoubleClick'
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ContextMenuStrip'
                                        Dock        = 'Fill'
                                        Events      = @{
                                            Opening    = [Scriptblock]{ # Event
                                                Invoke-EventTracer $this 'MenuEnter'
                                                $Nb = $Global:ControlHandler['DataGridView_ADAccounts'].SelectedRows.count
                                                $Global:ControlHandler['ContextMenuStrip_ADAccounts_Title'].text = "$Nb Accounts"
                                            }
                                        }
                                        Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{
                                            ControlType = 'ToolStripLabel'
                                            Name        = 'ContextMenuStrip_ADAccounts_Title'
                                            Text        = 'X Accounts'
                                        },
                                        @{ ControlType = 'ToolStripSeparator'
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'Delete'
                                            ShortcutKeyDisplayString = 'Supp.'
                                            Text        = 'Supprimer le compte'
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $Global:ControlHandler['DataGridView_ADAccounts'].SelectedRows | ForEach-Object {
                                                        try {
                                                            $_.Tag | Remove-ADUser -Confirm
                                                            $Global:ControlHandler['DataGridView_ADAccounts'].Rows.Remove($_)
                                                        } catch {
                                                            Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                                                        }
                                                    }
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        },
                                        @{ ControlType = 'ToolStripSeparator'
                                        },
                                        @{
                                            ControlType = 'ToolStripMenuItem'
                                            ShortcutKeys = 'Ctrl+D'
                                            ShortcutKeyDisplayString = 'Ctrl+D'
                                            Text        = "Dupliquer"
                                            Enabled     = $False
                                            Events      = @{
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this.Text 'Click'
                                                    $Global:ControlHandler['DataGridView_ADAccounts'].SelectedRows
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        }
                                    )
                                },
                                @{  HeaderText = 'NtAccountName'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'Display Name'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 160
                                },
                                @{  HeaderText = 'Email'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 160
                                },
                                @{  HeaderText = 'Type'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'Status'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 80
                                },
                                @{  HeaderText = 'Phone'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 110
                                },
                                @{  HeaderText = 'Description'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 240
                                },
                                @{  HeaderText = 'OU'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'Sid'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 45
                                },
                                @{  HeaderText = 'PasswordType'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'CreateDate'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'LastLogonDate'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'ExpireDate'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 120
                                },
                                @{  HeaderText = 'Groupes'
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    Width = 200
                                }
                            )
                        },
                        @{  ControlType = 'ProgressBar'
                            Name        = 'ProgressBar_DataGridView_ADAccounts'
                            Style       = 'Marquee'
                            Dock        = 'Top'
                            Height      = 5
                        }
                    )
                }
            )
        },
        @{  ControlType = 'Splitter'
            Dock        = 'Left'
        },
        @{  ControlType = 'Panel'
            Name        = 'PanelLeft'
            Dock        = 'left'
            Events      = @{
                Enter = [Scriptblock]{ # Event
                    Invoke-EventTracer $this 'Enter'
                    # $Global:ControlHandler['PanelLeft'].Width = 640
                    # $Global:ControlHandler['PanelRight'].Width = 100
                    # $Global:ControlHandler['Logs'].Height = 58
                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'GroupBox'
                    Name        = 'TargetForest'
                    Text        = 'Selected Forest'
                    Dock        = 'Fill'
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'TabControl'
                            Name        = 'TabsSelectedForest'
                            Dock        = 'Fill'
                            Events      = @{
                                SelectedIndexChanged = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'SelectedIndexChanged'
                                    Invoke-EventTracer $Tabs.SelectedTab.Name $EventName
                                    switch ($this.SelectedTab.Name) {
                                        'Servers' {
                                            $Global:mantis.SelectedDomain.Servers.Get()
                                        }
                                        'Groups' {
                                            $Global:mantis.SelectedDomain.Groups.Get()
                                        }
                                        'DFS' {
                                            if ($Mantis.SelectedDomain.DFS.Root -notlike $Global:ControlHandler['TreeDFS'].TopNode.FullPath) {
                                                Start-Loading 'TreeDFS'
                                                Update-MantisDFS
                                            }
                                        }
                                        'Trash' {
                                            $Global:mantis.SelectedDomain.Trash.Get()
                                        }
                                        default {}
                                    }
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                @{  ControlType = 'TabPage'
                                    Name        = 'Servers'
                                    Text        = 'Servers'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType      = 'ListView'
                                            Name             = 'ListServers'
                                            Dock             = 'Fill'
                                            FullRowSelect    = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick' # left and right but not on empty area
                                                    try {
                                                        $cred = (get-CredentialByRegistry -ntAccountName (whoami.exe))
                                                        Open-RdSession -ComputerName $this.SelectedItems.text -Credential $cred| Out-Null
                                                    } catch {
                                                        Write-LogStep -prefix "L.$($_.InvocationInfo.ScriptLineNumber)" "", $_ error
                                                    }
                                                }
                                                ItemSelectionChanged  = [Scriptblock]{ # Event
                                                    if($_.IsSelected){
                                                        $Global:LVSrvChange = $_.item.Text
                                                    } else { 
                                                        $Global:LVSrvChange = $_.item.Text
                                                    } 
                                                    $Global:ControlHandler['DataGridView_Sessions'].Visible = $False
                                                }
                                                KeyDown = [Scriptblock]{ # Event
                                                    $Global:LVSrvKeyDown = $_
                                                }
                                                KeyUp = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'KeyUp'
                                                    $Global:LVSrvKeyDown = $null
                                                    if ($Global:LVSrvChange){
                                                        Set-SelectedRDServers
                                                    }
                                                }
                                                MouseUp = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'MouseUp'
                                                    if ($Global:LVSrvChange -and !($Global:LVSrvKeyDown.Control -or $Global:LVSrvKeyDown.Shift)){
                                                        Set-SelectedRDServers
                                                    }
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'DNSHostName'
                                                    Width        = 160
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'IP'
                                                    Width        = 90
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'OS'
                                                    Width        = 200
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Install'
                                                    Width        = 100
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'RDP Version'
                                                    Width        = 30
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'DN'
                                                    Width        = 0
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_ListServers'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'Groups'
                                    Text        = 'Groups'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType      = 'ListView'
                                            Name             = 'ListGroups'
                                            Dock             = 'Fill'
                                            # Activation       = 'OneClick'
                                            FullRowSelect    = $True
                                            # HoverSelection   = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'Click' # left and right but not on empty area
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                    $this.SelectedItems | Write-Object -PassThru
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Name'
                                                    Width        = 160
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = '#'
                                                    Width        = 50
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Imbriqué'
                                                    Width        = 150
                                                },
                                                @{  ControlType = 'ContextMenuStrip'
                                                    # Name        = 'ContextMenuStrip_MainContext'
                                                    Events      = @{
                                                        Enter    = [Scriptblock]{ # Event
                                                            Invoke-EventTracer $this 'Enter'
                                                        }
                                                    }
                                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                        @{
                                                            ControlType = 'ToolStripMenuItem'
                                                            Text        = 'Supprimer !'
                                                            # Enabled     = $False
                                                            Events      = @{
                                                                Click    = [Scriptblock]{ # Event
                                                                    Invoke-EventTracer $this 'Click'
                                                                    $Global:ControlHandler['ListGroups'].SelectedItems | ForEach-Object {
                                                                        $_.Tag | Remove-ADGroup -Confirm:$False
                                                                        $_.remove()
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    )
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_ListGroups'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'DFS'
                                    Text        = 'DFS'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'TreeView'
                                            Name        = 'TreeDFS'
                                            Dock        = 'Fill'
                                            ShowNodeToolTips = $true
                                            Events      = @{
                                                AfterExpand    = [Scriptblock]{ # Event
                                                    $item = $_
                                                    Invoke-EventTracer $this "'BeforeExpand' $($item.Node.Text)"
                                                    # $item.Node.FirstNode | Write-Object -PassThru
                                                    if ($item.Node.FirstNode.Tag -eq '-') {
                                                        $item.Node.fullPath | get-childitem -Directory -Force -ea 0 | ForEach-Object {
                                                            [PSCustomObject]@{
                                                                Name = $_.Name
                                                                Handler = $_.FullName
                                                                ToolTipText = "$Prefix$($_.FullName)"
                                                                ForeColor =[system.Drawing.Color]::DarkCyan
                                                            }
                                                        } | Update-TreeView -treeNode $item.Node -Clear -ChildrenScriptBlock {
                                                            [PSCustomObject]@{
                                                                Name = '-'
                                                                Handler = '-'
                                                                ForeColor =[system.Drawing.Color]::LightGray
                                                            }
                                                        }
                                                    }
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                    $this | Write-Object -PassThru
                                                    if ($this.SelectedNode.fullPath) {
                                                        Start-Process $this.SelectedNode.fullPath -ea 0
                                                    }
                                                    # $_.Cancel = $true
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_TreeDFS'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    # Name        = 'TabPage'
                                    Text        = 'Trash'
                                    Dock        = 'Fill'
                                    Events      = @{
                                    }
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ListView'
                                            Name             = 'ListTrash'
                                            Dock             = 'Fill'
                                            # Activation       = 'OneClick'
                                            FullRowSelect    = $True
                                            # HoverSelection   = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'Click' # left and right but not on empty area
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                    $this.SelectedItems | Write-Object -PassThru
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Name'
                                                    Width        = 150
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'FullName'
                                                    Width        = 0
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Deleted On'
                                                    Width        = 140
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'LastKnownPath'
                                                    Width        = 300
                                                },
                                                @{  ControlType = 'ContextMenuStrip'
                                                    # Name        = 'ContextMenuStrip_MainContext'
                                                    Events      = @{
                                                        Enter    = [Scriptblock]{ # Event
                                                            Invoke-EventTracer $this 'Enter'
                                                        }
                                                    }
                                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                        @{
                                                            ControlType = 'ToolStripMenuItem'
                                                            ShortcutKeys = 'Ctrl+Z'
                                                            ShortcutKeyDisplayString = 'Ctrl+Z'
                                                            Text        = 'Restaurer'
                                                            Events      = @{
                                                                Click    = [Scriptblock]{ # Event
                                                                    Invoke-EventTracer $this 'Click'
                                                                    $Global:ControlHandler['ListTrash'].SelectedItems | ForEach-Object {
                                                                        $_.Tag | Restore-ADObject -PassThru | ForEach-Object { $_ | Set-ADUser -Description "Restaured on $(get-date) by $(whoami.exe) ($_)" }
                                                                        $_.remove()
                                                                    }
                                                                }
                                                            }
                                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                            )
                                                        },
                                                        @{
                                                            ControlType = 'ToolStripSeparator'
                                                        },
                                                        @{
                                                            ControlType = 'ToolStripMenuItem'
                                                            Text        = 'Supprimer definitivement !'
                                                            # Enabled     = $False
                                                            Events      = @{
                                                                Click    = [Scriptblock]{ # Event
                                                                    Invoke-EventTracer $this 'Click'
                                                                    $Global:ControlHandler['ListTrash'].SelectedItems | ForEach-Object {
                                                                        $_.Tag | Remove-ADObject -IncludeDeletedObjects -Confirm
                                                                        $_.remove()
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    )
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_ListTrash'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                        }
                                    )
                                }
                            )
                        }
                    )
                },
                @{  ControlType = 'GroupBox'
                    Text        = 'TreeForest'
                    Dock        = 'top'
                    Height      = 180
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'TreeView'
                            Name        = 'TreeForest'
                            HideSelection = $False
                            ShowNodeToolTips = $true
                            Dock        = 'Fill'
                            Events      = @{
                                BeforeExpand    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'BeforeExpand'
                                }
                                DoubleClick    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'DoubleClick'
                                }
                                AfterSelect =[Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'AfterSelect'
                                    Get-Job | Where-Object {
                                        $_.Name -match "^Mantis_(\w|-)+_.+"
                                    } | Remove-Job -Force
                                    $Global:mantis.Domain($this.SelectedNode.Text)
                                    $Global:SequenceStart.Restart()
                                    $Global:ControlHandler['TargetForest'].Text = "Forest : [$($this.SelectedNode.Text)]"
                                    
                                    # $tab = $($Global:ControlHandler['TabsSelectedForest'].SelectedTab.Name)
                                    # $Global:ControlHandler['TabsSelectedForest'].DeselectTab($tab)
                                    # $Global:ControlHandler['TabsSelectedForest'].SelectTab($tab)
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                            )
                        }
                    )
                }
            )
        },
        @{  ControlType = 'Splitter'
            Dock        = 'Right'
        },
        @{  ControlType = 'Panel'
            Name        = 'PanelRight'
            Dock        = 'Right'
            Events      = @{
                Enter = [Scriptblock]{ # Event
                    Invoke-EventTracer $this 'Enter'
                    # $Global:ControlHandler['PanelLeft'].Width = 220
                    # $Global:ControlHandler['PanelRight'].Width = 320
                    # $Global:ControlHandler['Logs'].Height = 58

                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'GroupBox'
                    Name        = 'ServerNameGbx'
                    Text        = 'Server'
                    Dock        = 'Fill'
                    Events      = @{
                        Enter = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'Enter'
                            # $Global:ControlHandler['UserNameGbx'].Height = [int]($Global:ControlHandler['PanelRight'].Height * 0.2)
                        }
                    }
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'TabControl'
                            Dock        = 'Fill'
                            Events      = @{
                                SelectedIndexChanged = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'SelectedIndexChanged'
                                    switch ($this.SelectedTab.Name) {
                                        'SrvProp' {}
                                        'SrvPrinters' {}
                                        'SrvReg' {}
                                        default {}
                                    }
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                @{  ControlType = 'TabPage'
                                    Name        = 'SrvProp'
                                    Text        = 'Properties'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ListView'
                                            Name             = 'LVServerADProp'
                                            Dock             = 'Fill'
                                            # Activation       = 'OneClick'
                                            FullRowSelect    = $True
                                            # HoverSelection   = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{  Text        = 'Prop'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 120
                                                },
                                                @{  Text        = 'Value'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 220
                                                },
                                                @{  Text        = 'Complements'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 0
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_LVServerADProp'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                            visible     = $False
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'SrvPrinters'
                                    Text        = 'Printers'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'SrvReg'
                                    Text        = 'Registre'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'SrvProfiles'
                                    Text        = 'Profiles'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                }
                            )
                        }
                    )
                },
                @{  ControlType = 'Splitter'
                    Dock        = 'Bottom'
                },
                @{  ControlType = 'GroupBox'
                    Name        = 'UserNameGbx'
                    Text        = 'User'
                    Dock        = 'Bottom'
                    Height      = 240
                    Events      = @{
                        Enter = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'Enter'
                            $This.Height = [int]($Global:ControlHandler['PanelRight'].Height * 0.8)
                        }
                    }
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'TabControl'
                            Dock        = 'Fill'
                            Events      = @{
                                SelectedIndexChanged = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'SelectedIndexChanged'
                                    switch ($this.SelectedTab.Name) {
                                        'UserADProp'{

                                        }
                                        'UserGroups'{

                                        }
                                        'UsersPrinters'{

                                        }
                                        'UsersPaths'{

                                        }
                                        'UsersReg' {

                                        }
                                        default {}
                                    }
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserADProp'
                                    Text        = 'Properties'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ListView'
                                            Name             = 'LVUserADProp'
                                            Dock             = 'Fill'
                                            # Activation       = 'OneClick'
                                            FullRowSelect    = $True
                                            # HoverSelection   = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{  Text        = 'Prop'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 120
                                                },
                                                @{  Text        = 'Value'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 220
                                                },
                                                @{  Text        = 'Complements'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 0
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_LVUserADProp'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                            visible     = $False
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserGroups'
                                    Text        = 'Groups'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ListView'
                                            Name             = 'ListView_LVUserGroups'
                                            Dock             = 'Fill'
                                            # Activation       = 'OneClick'
                                            FullRowSelect    = $True
                                            # HoverSelection   = $True
                                            ShowGroups       = $True
                                            ShowItemToolTips = $True
                                            View             = 'Details'
                                            Events      = @{
                                                Enter    = @{ # Event
                                                    Type = 'Thread'
                                                    ScriptBlock = [Scriptblock]{
                                                        param($ControlHandler, $ControllerName, $EventName)
                                                        $Tabs = $ControlHandler['TabsRessources']
                                                        $ListView = $($ControlHandler["$($Tabs.SelectedTab.Name)_LV"])
                                                    }
                                                }
                                                ColumnClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'ColumnClick'
                                                    Set-ListViewSorted  -listView $this -column $_.Column
                                                }
                                                Click    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'Click' # left and right but not on empty area
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{  Text        = 'Col1'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 140
                                                },
                                                @{  Text        = 'Col2'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 120
                                                }
                                                @{  Text        = 'Col3'
                                                    ControlType = 'ColumnHeader'
                                                    Width        = 150
                                                }
                                            )
                                        },
                                        @{  ControlType = 'ProgressBar'
                                            Name        = 'ProgressBar_LVUserGroups'
                                            Style       = 'Marquee'
                                            Dock        = 'Top'
                                            Height      = 5
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserPrinters'
                                    Text        = 'Printers'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserPaths'
                                    Text        = 'Paths'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserReg'
                                    Text        = 'Registre'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                    )
                                }
                            )
                        }
                    )
                }
            )
        },
        @{  ControlType = 'MenuStrip' # Main Top menu need last child
            Dock        = 'Top'
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{
                    ControlType = 'ToolStripMenuItem'
                    # Name        = 'ToolStripMenuItem'
                    Text        = 'Menu'
                    Events      = @{
                        Click    = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'Click'
                        }
                    }
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'ToolStripMenuItem'
                            # Name        = 'ToolStripMenuItem'
                            ShortcutKeys = 'Ctrl+S'
                            ShortcutKeyDisplayString = 'Ctrl+S'
                            Text        = 'Settings'
                            Events      = @{
                                Click    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Click'
                                }
                            }
                        }
                    )
                },
                @{
                    ControlType = 'ToolStripMenuItem'
                    # Name        = 'ToolStripMenuItem'
                    Text        = 'Aide'
                    Events      = @{
                        Click    = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'Click'
                        }
                    }
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'ToolStripMenuItem'
                            # Name        = 'ToolStripMenuItem'
                            ShortcutKeys = 'F1'
                            ShortcutKeyDisplayString = 'F1'
                            Text        = 'Aide en ligne'
                            Events      = @{
                                Click    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Click'
                                    $Me = Import-Module RDS.Mantis -Force -PassThru
                                    Start-Process $me.HelpInfoURI
                                }
                            }
                        },
                        @{  ControlType = 'ToolStripMenuItem'
                            # Name        = 'ToolStripMenuItem'
                            Text        = 'A Propos'
                            Events      = @{
                                Click    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'Click'
                                    $Me = Import-Module RDS.Mantis -Force -PassThru
                                    PopUp-Box $me.Name -Message "Version $($Me.Version.toString())`nBy $($me.Author)`n`n$($me.Description)" -Buttons OK -Icon Information
                                }
                            }
                        }
                    )
                }
            )
        },
        @{  ControlType = 'Splitter'
            Dock        = 'Bottom'
        },
        @{  ControlType = 'GroupBox'
            Name        = 'Logs'
            Text        = 'Logs'
            Height      = 128
            Dock        = 'Bottom'
            Events      = @{
                Enter = [Scriptblock]{ # Event
                    Invoke-EventTracer $this 'SelectedIndexChanged'
                    # $Global:ControlHandler['Logs'].Height = 240
                    # $Global:ControlHandler['UserNameGbx'].Height = [int]($Global:ControlHandler['PanelRight'].Height * 0.2)
                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'RichTextBox'
                    Name        = 'RichTextBox_Logs'
                    Dock        = 'fill'
                    Events      = @{
                        DoubleClick    = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'DoubleClick'
                        }
                    }
                }
            )
        }
    )
}