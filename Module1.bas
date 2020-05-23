Attribute VB_Name = "Module1"
Public Function GetWMI_Board()
    Dim List, LQry, Msg, Obj

    On Local Error Resume Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}")
    
    Set LQry = List.ExecQuery("SELECT * FROM Win32_BaseBoard")
    For Each Obj In LQry
        Msg = Msg & PadString(" Board Manufacturer", 30) & " : " & Obj.Manufacturer & vbCrLf
        Msg = Msg & PadString(" Board Serial No", 30) & " : " & Obj.SerialNumber & vbCrLf
    Next
    
    Set LQry = List.ExecQuery("SELECT * FROM Win32_BIOS")
    For Each Obj In LQry
        Msg = Msg & PadString(" BIOS Name", 30) & " : " & Obj.Caption & vbCrLf
        Msg = Msg & PadString(" BIOS Maker", 30) & " : " & Obj.Manufacturer & vbCrLf
        Msg = Msg & PadString(" BIOS Serial No", 30) & " : " & Obj.SerialNumber & vbCrLf
        Msg = Msg & PadString(" BIOS Release Date", 30) & " : " & DateParse(Obj.ReleaseDate) & vbCrLf
    Next

    Msg = Msg & vbCrLf
    
    Set LQry = List.ExecQuery("SELECT * FROM Win32_ComputerSystem")
    For Each Obj In LQry
        Msg = Msg & PadString(" PC Model", 30) & " : " & Obj.Model & vbCrLf
        Msg = Msg & PadString(" PC Memmory", 30) & " : " & FormatSize(Obj.TotalPhysicalMemory) & vbCrLf
        Msg = Msg & PadString(" PC # of Processors", 30) & " : " & Obj.NumberOfProcessors & vbCrLf
        Msg = Msg & PadString(" PC # of Logical Processors", 30) & " : " & Obj.NumberOfLogicalProcessors & vbCrLf

        Msg = Msg & PadString(" Logon User", 30) & " : " & Obj.UserName & vbCrLf
        Msg = Msg & PadString(" Domain", 30) & " : " & Obj.Domain & vbCrLf
    Next


    Msg = Msg & vbCrLf

    Set LQry = List.ExecQuery("SELECT * FROM Win32_OperatingSystem")
    For Each Obj In LQry
        Msg = Msg & PadString(" OS Name", 30) & " : " & Obj.Caption & vbCrLf
        Msg = Msg & PadString(" OS Type", 30) & " : " & Obj.OSArchitecture & vbCrLf
        Msg = Msg & PadString(" OS Path", 30) & " : " & Obj.WindowsDirectory & vbCrLf
        Msg = Msg & PadString(" Hostname", 30) & " : " & Obj.CSName & vbCrLf
        Msg = Msg & PadString(" Date Installed", 30) & " : " & DateParse(Obj.InstallDate) & vbCrLf
        Msg = Msg & PadString(" Date Uptime", 30) & " : " & DateParse(Obj.LastBootUpTime) & vbCrLf
        
    Next
    
   
GetWMI_Board = Msg & vbCrLf

End Function

Public Function GetWMI_Accounts()
    Dim List, LQry, Msg, Obj, sThisPC
    Dim cnt
    
    cnt = 1

    On Local Error Resume Next
    
    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}")
    Set LQry = List.ExecQuery("SELECT * FROM Win32_UserAccount WHERE Domain = '" & Environ("computername") & "'")
    
    For Each Obj In LQry
        Msg = Msg & PadString(" User " & cnt, 30) & " : " & Obj.Caption & vbCrLf
        cnt = cnt + 1
    Next

GetWMI_Accounts = Msg & vbCrLf

End Function

Public Function GetWMI_Graphics()
    Dim List, Msg, Obj
    Dim cnt
   
    On Local Error Resume Next
    
    cnt = 1

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_VideoController")
    For Each Obj In List
    
        Msg = Msg & PadString(" Display " & cnt, 30) & " : " & Obj.Caption & vbCrLf
        Msg = Msg & PadString("   Resolution", 30) & " : " & Obj.CurrentHorizontalResolution & " x " & Obj.CurrentVerticalResolution & vbCrLf
        Msg = Msg & PadString("   Refresh Rate", 30) & " : " & Obj.CurrentRefreshRate & vbCrLf
        'Msg = Msg & PadString("   Memory", 30) & " : " & Obj.AdapterRAM & vbCrLf
        Msg = Msg & PadString("   Driver Version", 30) & " : " & Obj.DriverVersion & vbCrLf
        Msg = Msg & PadString("   Driver Date", 30) & " : " & DateParse(Obj.DriverDate) & vbCrLf & vbCrLf
        cnt = cnt + 1
    Next

GetWMI_Graphics = Msg & vbCrLf

End Function

Public Function GetWMI_Printers()
    Dim List, Msg, Obj
    Dim cnt
   
    On Local Error Resume Next
    
    cnt = 1

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_Printer")
    For Each Obj In List
        Msg = Msg & PadString(" Printer " & cnt, 30) & " : " & Obj.DriverName & vbCrLf
        Msg = Msg & PadString("   Port ", 30) & " : " & Obj.PortName & vbCrLf
        Msg = Msg & PadString("   Shared ", 30) & " : " & Obj.Shared & vbCrLf & vbCrLf
        cnt = cnt + 1
    Next

GetWMI_Printers = Msg & vbCrLf

End Function
Public Function GetWMI_NetAdapters()
    Dim List, LQry, Msg, Obj
    Dim cnt
    
    On Local Error Resume Next

    cnt = 1

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}")
    Set LQry = List.ExecQuery("Select * from Win32_NetworkAdapter WHERE PhysicalAdapter = 'True'")
    
    For Each Obj In LQry
        Msg = Msg & PadString(" Network Adapter " & cnt, 30) & " : " & Obj.ProductName & vbCrLf
        Msg = Msg & PadString("   Maker", 30) & " : " & Obj.Manufacturer & vbCrLf
        Msg = Msg & PadString("   MAC Address", 30) & " : " & Obj.MACAddress & vbCrLf & vbCrLf
        cnt = cnt + 1
    Next


    Msg = Msg & "==============================================================================================================" & _
          vbCrLf & vbCrLf

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_IP4RouteTable")

    Msg = Msg & PadString(" Network Destination", 25) & _
                        PadString("Netmask", 25) & _
                        PadString("Gateway", 25) & _
                        PadString("Metric", 25) & vbCrLf
                        
    For Each Obj In List
        
        Msg = Msg & PadString(" " & Obj.Destination, 25) & _
                    PadString(Obj.Mask, 25) & _
                    PadString(Obj.NextHop, 25) & _
                    PadString(Obj.Metric1, 25) & vbCrLf
       
    Next

GetWMI_NetAdapters = Msg & vbCrLf

End Function


Public Function GetWMI_Drives()
    Dim List, Msg, Obj
    Dim cnt
    
    On Local Error Resume Next

    cnt = 1

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_DiskDrive")
    For Each Obj In List
        Msg = Msg & PadString(" Physical Disk " & cnt, 30) & " : " & Obj.Caption & vbCrLf
        Msg = Msg & PadString("   Size ", 30) & " : " & FormatSize(Obj.Size) & vbCrLf
        Msg = Msg & PadString("   Partitions ", 30) & " : " & Obj.Partitions & vbCrLf & vbCrLf
        cnt = cnt + 1
    Next

    Msg = Msg & "==============================================================================================================" & _
          vbCrLf & vbCrLf

    Set List = GetObject("winmgmts:{impersonationLevel=impersonate}").InstancesOf("Win32_LogicalDisk")
    For Each Obj In List
    
        If Obj.DriveType <= 3 Or Obj.DriveType = 5 Then ' Disk and CD Drive
            Msg = Msg & PadString(" Drive " & Obj.Caption, 30) & " : " & Obj.VolumeName & vbCrLf
        End If
        
        If Obj.DriveType = 4 Or Obj.DriveType > 5 Then ' Network Drive
            Msg = Msg & PadString(" Drive " & Obj.Caption, 30) & " : " & Obj.ProviderName & vbCrLf
        End If

        Msg = Msg & PadString("   Type", 30) & " : " & Obj.Description & vbCrLf
        
        If Obj.Size <> "" Then
            Msg = Msg & PadString("   File System", 30) & " : " & Obj.FileSystem & vbCrLf
            Msg = Msg & PadString("   Total Used", 30) & " : " & FormatSize(Obj.Size - Obj.FreeSpace) & vbCrLf
            Msg = Msg & PadString("   Total Free", 30) & " : " & FormatSize(Obj.FreeSpace) & vbCrLf
            Msg = Msg & PadString("   Total Capacity", 30) & " : " & FormatSize(Obj.Size) & vbCrLf
        End If
        
        Msg = Msg & vbCrLf
    Next

GetWMI_Drives = Msg & vbCrLf

End Function

Function GetAddRemove()
    Dim aTmp
    Dim cnt, oReg, sBaseKey, iRC, aSubKeys
    
    Const HKCU = &H80000001 'HKEY_CURRENT_USER
    Const HKLM = &H80000002 'HKEY_LOCAL_MACHINE
    
    On Local Error Resume Next
    
    Set oReg = GetObject("winmgmts:{impersonationLevel=impersonate}!\\" & Environ("computername") & "/root/default:StdRegProv")
    sBaseKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\"

    Dim sKey, sValue, sTmp, sVersion, sDateValue, sYr, sMth, sDay

    '---------------------------------------------------------------------------------------------------------------
    ' LOCAL MACHINE
    '---------------------------------------------------------------------------------------------------------------
    iRC = oReg.EnumKey(HKLM, sBaseKey, aSubKeys)
    
    For Each sKey In aSubKeys
        iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, "DisplayName", sValue)
      
        If iRC <> 0 Then
            oReg.GetStringValue HKLM, sBaseKey & sKey, "QuietDisplayName", sValue
        End If
      
        If sValue <> "" Then
            iRC = oReg.GetStringValue(HKLM, sBaseKey & sKey, "DisplayVersion", sVersion)
            If sVersion <> "" Then
              sValue = sValue & ", Ver: " & sVersion
            Else
              sValue = sValue
            End If
            sTmp = sTmp & "  " & Trim(sValue) & vbCrLf
            
          cnt = cnt + 1
        End If
    Next
    
    '---------------------------------------------------------------------------------------------------------------
    ' CURRENT USER
    '---------------------------------------------------------------------------------------------------------------
    iRC = oReg.EnumKey(HKCU, sBaseKey, aSubKeys)
    
    For Each sKey In aSubKeys
        iRC = oReg.GetStringValue(HKCU, sBaseKey & sKey, "DisplayName", sValue)
      
        If iRC <> 0 Then
            oReg.GetStringValue HKCU, sBaseKey & sKey, "QuietDisplayName", sValue
        End If
      
        If sValue <> "" Then
            iRC = oReg.GetStringValue(HKCU, sBaseKey & sKey, "DisplayVersion", sVersion)
            If sVersion <> "" Then
              sValue = sValue & ", Ver: " & sVersion
            Else
              sValue = sValue
            End If
            sTmp = sTmp & "  " & Trim(sValue) & vbCrLf
            
          cnt = cnt + 1
        End If
    Next
    
    '---------------------------------------------------------------------------------------------------------------
    sTmp = AutoSort(sTmp)

    GetAddRemove = sTmp
End Function


