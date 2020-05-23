VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00000000&
   Caption         =   "PC Audit"
   ClientHeight    =   7485
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   11550
   BeginProperty Font 
      Name            =   "Lucida Console"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11550
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   1920
      Top             =   5040
   End
   Begin VB.TextBox txtStatus 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3135
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   825
      Width           =   5655
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public ThisPC As String
Private Sub Form_Resize()
    txtStatus.Top = 0
    txtStatus.Left = 0
    txtStatus.Height = Me.ScaleHeight   '- (txtStatus.Top + 650)
    txtStatus.Width = Me.ScaleWidth  '- (txtStatus.Left + 300)
End Sub

Private Sub Form_Load()
    ThisPC = Environ("computername")
    Me.Caption = "Ayanix PC Audit v" & App.Major & "." & App.Minor & "." & App.Revision
    Form_Resize
End Sub

Private Sub StartWMI()
    Dim Msg, SaveFile
  
    On Local Error Resume Next

    SaveFile = ThisPC & "_" & Format(Now(), "ddMMyyyyHHmmss") & ".txt"
      
    Msg = Msg & TitleSeparator("GENERAL INFORMATION")
    Msg = Msg & GetWMI_Board()
    
    Msg = Msg & TitleSeparator("DISPLAY ADAPTER")
    Msg = Msg & GetWMI_Graphics()

    Msg = Msg & TitleSeparator("ALL NETWORK ADAPTERS")
    Msg = Msg & GetWMI_NetAdapters()
   
    Msg = Msg & TitleSeparator("ALL STORAGE DRIVES")
    Msg = Msg & GetWMI_Drives()
    
    Msg = Msg & TitleSeparator("ALL LOCAL ACCOUNTS")
    Msg = Msg & GetWMI_Accounts()
    
    Msg = Msg & TitleSeparator("ALL PRINTERS")
    Msg = Msg & GetWMI_Printers()
    
    Msg = Msg & TitleSeparator("ALL INSTALLED SOFTWARE")
    Msg = Msg & GetAddRemove()

    txtStatus.Text = Msg

    On Error GoTo Er

    Open App.Path & "\" & SaveFile For Output As #1
    Print #1, Msg
    Close #1
        
    'Me.Caption = ThisPC & " " & SaveFile
        
    MsgBox "File " & SaveFile & " successfully saved.", vbInformation, "Saved File"
    
    Exit Sub
Er:
    MsgBox "File " & SaveFile & " error saving file " & Err.Description, vbCritical, "Saved File Error"
    
End Sub



Private Sub Timer1_Timer()
    txtStatus.Text = "Starting " & ThisPC & " scan..." & vbCrLf & vbCrLf & _
                     "Note : This App may freeze for a moment while collecting data." & vbCrLf & vbCrLf & _
                     "Please wait....       "
    StartWMI
    
    Timer1.Enabled = False
End Sub
