Attribute VB_Name = "Tools"
Function AutoSort(sTmp)
  Dim aTmp, i, j, temp
  aTmp = Split(sTmp, vbCrLf)
  For i = UBound(aTmp) - 1 To 0 Step -1
    For j = 0 To i - 1
      If LCase(aTmp(j)) > LCase(aTmp(j + 1)) Then
        temp = aTmp(j + 1)
        aTmp(j + 1) = aTmp(j)
        aTmp(j) = temp
      End If
    Next
  Next
  AutoSort = Join(aTmp, vbCrLf)
End Function

Public Function PadString(strSource As String, lPadLen As Long, Optional PadChar As String = " ") As String
    PadString = String(lPadLen, PadChar)
    LSet PadString = strSource
End Function

Public Function FormatSize(ByVal Size As Currency) As String
    Const Kilobyte As Currency = 1024@
    Const TenK As Currency = 10240@
    Const HundredK As Currency = 102400@
    Const ThousandK As Currency = 1024000@
    Const Megabyte As Currency = 1048576@
    Const TenMeg As Currency = 10485760@
    Const HundredMeg As Currency = 104857600@
    Const ThousandMeg As Currency = 1048576000@
    Const Gigabyte As Currency = 1073741824@
    Const TenGig As Currency = 10737418240@
    Const HundredGig As Currency = 107374182400@
    Const ThousandGig As Currency = 1073741824000@
    Const Terabyte As Currency = 1099511627776@
    
    Select Case Size
        Case Is < Kilobyte: FormatSize = Int(Size) & " bytes"
        Case Is < TenK: FormatSize = Format(Size / Kilobyte, "0.00") & " KB"
        Case Is < HundredK: FormatSize = Format(Size / Kilobyte, "0.0") & " KB"
        Case Is < ThousandK: FormatSize = Int(Size / Kilobyte) & " KB"
        Case Is < TenMeg: FormatSize = Format(Size / Megabyte, "0.00") & " MB"
        Case Is < HundredMeg: FormatSize = Format(Size / Megabyte, "0.0") & " MB"
        Case Is < ThousandMeg: FormatSize = Int(Size / Megabyte) & " MB"
        Case Is < TenGig: FormatSize = Format(VBA.Round(Size / Gigabyte), "0.0") & " GB"
        Case Is < HundredGig: FormatSize = Format(Size / Gigabyte, "0.0") & " GB"
        Case Is < ThousandGig: FormatSize = Int(Size / Gigabyte) & " GB"
        Case Else: FormatSize = Format(Size / Terabyte, "0.00") & " TB"
    End Select

End Function

Public Function TitleSeparator(Title As String)
    Dim sTmp

    sTmp = "==============================================================================================================" & vbCrLf
    sTmp = sTmp & " " & Title & vbCrLf
    sTmp = sTmp & "==============================================================================================================" & vbCrLf & vbCrLf

TitleSeparator = sTmp
End Function

Public Function DateParse(sDate)
    Dim sTmp, sTmp2, sM, sD, sY, sTime

    sY = Left$(sDate, 4)
    sM = Mid$(sDate, 5, 2)
    sD = Mid$(sDate, 7, 2)
    
    sTime = Mid$(sDate, 9, 2) & ":" & Mid$(sDate, 11, 2) & ":" & Mid$(sDate, 13, 2)
    sTmp = sM & "/" & sD & "/" & sY & " " & sTime
        
DateParse = sTmp
End Function


