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
            Invoke-EventTracer $this 'KeyDown'
            if ($_.KeyCode -eq 'Escape') {
                $this.Close()
            }
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
                    $Global:ControlHandler['PanelLeft'].Width = 160
                    $Global:ControlHandler['PanelRight'].Width = 100
                    $Global:ControlHandler['Logs'].Height = 58
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
                    Text        = 'Ferme RDS'
                    Dock        = 'Fill'
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'DataGridView'
                            Name             = 'DataGridView_label'
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
                                }
                                DoubleClick    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'DoubleClick'
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ComputerName'
                                    Width = 140
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'SessionID'
                                    Width = 35
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'State'
                                    Width = 80
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Sid'
                                    Width = 45
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'UserAccount'
                                    Width = 150
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'IPAddress'
                                    Width = 125
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ClientName'
                                    Width = 130
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Protocole'
                                    Width = 75
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ClientBuildNumber'
                                    Width = 60
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'LoginTime'
                                    Width = 115
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ConnectTime'
                                    Width = 115
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'DisconnectTime'
                                    Width = 115
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Inactivite'
                                    Width = 70
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Screen'
                                    Width = 100
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Process'
                                    Width = 230
                                }
                            )
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
                                }
                                DoubleClick    = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'DoubleClick'
                                }
                            }
                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'NtAccountName'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Display Name'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Sid'
                                    Width = 45
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Email'
                                    Width = 130
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Type'
                                    Width = 100
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'State'
                                    Width = 50
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Phone'
                                    Width = 110
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Description'
                                    Width = 60
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Facturation'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'PasswordType'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'CreateDate'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'LastLogonDate'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'ExpireDate'
                                    Width = 120
                                    
                                },
                                @{
                                    ControlType = 'DataGridViewTextBoxColumn'
                                    HeaderText = 'Groupes'
                                    Width = 200
                                }
                            )
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
                    $Global:ControlHandler['PanelLeft'].Width = 320
                    $Global:ControlHandler['PanelRight'].Width = 100
                    $Global:ControlHandler['Logs'].Height = 58
                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'GroupBox'
                    Name        = 'Target'
                    Text        = 'Forest'
                    Dock        = 'Fill'
                    Events      = @{}
                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                        @{  ControlType = 'TabControl'
                            Dock        = 'Fill'
                            Events      = @{
                                SelectedIndexChanged = [Scriptblock]{ # Event
                                    Invoke-EventTracer $this 'SelectedIndexChanged'
                                    Invoke-EventTracer $Tabs.SelectedTab.Name $EventName
                                    switch ($this.SelectedTab.Name) {
                                        'Servers' {
                                            $Global:mantis.SelectedDomain.Servers.Get()
                                        }
                                        'DFS' {
                                            [PSCustomObject]@{
                                                Name =  $mantis.SelectedDomain.DFS.Root
                                                Handler = $mantis.SelectedDomain.DFS.Root
                                                ToolTipText = $mantis.SelectedDomain.DFS.Root
                                            } | Update-TreeView -treeNode $Global:ControlHandler['TreeDFS'] -expand -Clear -Depth 1
                                            $mantis.SelectedDomain.DFS.Get() | Update-MantisDFS
                                        }
                                        'Groups' {
                                            $Global:mantis.SelectedDomain.Groups.Get()
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
                                        @{  ControlType = 'ListView'
                                            Name             = 'ListServers'
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
                                                    Width        = 100
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'OS'
                                                    Width        = 240
                                                }
                                            )
                                        }
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'Groups'
                                    Text        = 'Groups'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                        @{  ControlType = 'ListView'
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
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Name'
                                                    Width        = 140
                                                },
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Col2'
                                                    Width        = 120
                                                }
                                                @{
                                                    ControlType = 'ColumnHeader'
                                                    Text        = 'Col3'
                                                    Width        = 150
                                                }
                                            )
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
                                                BeforeExpand    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'BeforeExpand'
                                                    Get-ChildItem $this.Tag | ForEach-Object {
                                                        @{
                                                            Name = $_.name
                                                            Handler = $_.FullName
                                                            ToolTipText = $_.FullName
                                                            ForeColor = [System.Drawing.Color]::DarkBlue
                                                        } | Update-TreeView -treeNode $this -Clear -Depth 1
                                                    }
                                                }
                                                DoubleClick    = [Scriptblock]{ # Event
                                                    Invoke-EventTracer $this 'DoubleClick'
                                                }
                                            }
                                            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                                            )
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
                                    $Global:mantis.Domain($this.SelectedNode.Text)
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
                    $Global:ControlHandler['PanelLeft'].Width = 160
                    $Global:ControlHandler['PanelRight'].Width = 320
                    $Global:ControlHandler['Logs'].Height = 58
                }
            }
            Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
                @{  ControlType = 'GroupBox'
                    Name        = 'ServerName'
                    Text        = 'Server'
                    Dock        = 'Fill'
                    Events      = @{
                        Enter = [Scriptblock]{ # Event
                            Invoke-EventTracer $this 'Enter'
                            $Global:ControlHandler['UserName'].Height = [int]($Global:ControlHandler['PanelRight'].Height * 0.2)
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
                                }
                            )
                        }
                    )
                },
                @{  ControlType = 'Splitter'
                    Dock        = 'Bottom'
                },
                @{  ControlType = 'GroupBox'
                    Name        = 'UserName'
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
                                    )
                                },
                                @{  ControlType = 'TabPage'
                                    Name        = 'UserGroups'
                                    Text        = 'Groups'
                                    Dock        = 'Fill'
                                    Events      = @{}
                                    Childrens   = @( # FirstControl need {Dock = 'Fill'} but the following will be [Top, Bottom, Left, Right]
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
                    $Global:ControlHandler['Logs'].Height = 240
                    $Global:ControlHandler['UserName'].Height = [int]($Global:ControlHandler['PanelRight'].Height * 0.2)
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