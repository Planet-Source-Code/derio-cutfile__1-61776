VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "Comdlg32.ocx"
Begin VB.Form FMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Cut Files"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5025
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5025
   StartUpPosition =   3  'Windows Default
   Begin VB.Timer tmrSendEmail 
      Enabled         =   0   'False
      Interval        =   5000
      Left            =   120
      Top             =   2400
   End
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   60
      Top             =   1740
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Frame freFile 
      Height          =   3555
      Left            =   660
      TabIndex        =   1
      Top             =   120
      Width           =   4215
      Begin VB.CheckBox chkSendTo 
         Caption         =   "Send to"
         Height          =   255
         Left            =   300
         TabIndex        =   13
         Top             =   1380
         Width           =   975
      End
      Begin VB.TextBox txtAddress 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1440
         Locked          =   -1  'True
         TabIndex        =   12
         Top             =   1320
         Width           =   2115
      End
      Begin VB.Frame Frame2 
         Height          =   460
         Left            =   120
         TabIndex        =   11
         Top             =   3000
         Width           =   3945
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   15
            Left            =   3720
            Top             =   180
            Visible         =   0   'False
            Width           =   90
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   14
            Left            =   3480
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   13
            Left            =   3240
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   12
            Left            =   3000
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   11
            Left            =   2760
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   10
            Left            =   2520
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   9
            Left            =   2280
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   8
            Left            =   2040
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   7
            Left            =   1800
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   6
            Left            =   1560
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   5
            Left            =   1320
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   4
            Left            =   1080
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   3
            Left            =   840
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   2
            Left            =   600
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   1
            Left            =   360
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
         Begin VB.Shape shpProgress 
            BackColor       =   &H80000010&
            BackStyle       =   1  'Opaque
            Height          =   195
            Index           =   0
            Left            =   120
            Top             =   180
            Visible         =   0   'False
            Width           =   195
         End
      End
      Begin VB.TextBox txtFileSize 
         Alignment       =   1  'Right Justify
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   9
         Top             =   2460
         Width           =   1395
      End
      Begin VB.ComboBox cboFileSize 
         Height          =   315
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   2100
         Width           =   2355
      End
      Begin VB.TextBox txtOutput 
         BackColor       =   &H8000000F&
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   960
         Width           =   2355
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "..."
         Height          =   285
         Left            =   3540
         TabIndex        =   4
         Top             =   300
         Width           =   435
      End
      Begin VB.TextBox txtSource 
         Height          =   285
         Left            =   1200
         Locked          =   -1  'True
         TabIndex        =   3
         Top             =   300
         Width           =   2355
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "byte(s)"
         Height          =   195
         Index           =   3
         Left            =   2640
         TabIndex        =   10
         Top             =   2520
         Width           =   465
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "File Size"
         Height          =   195
         Index           =   2
         Left            =   300
         TabIndex        =   7
         Top             =   2160
         Width           =   585
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Output File"
         Height          =   195
         Index           =   1
         Left            =   300
         TabIndex        =   5
         Top             =   1020
         Width           =   765
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Source File"
         Height          =   195
         Index           =   0
         Left            =   300
         TabIndex        =   2
         Top             =   360
         Width           =   795
      End
   End
   Begin VB.CommandButton cmdExecute 
      Caption         =   "Execute"
      Height          =   375
      Left            =   3720
      TabIndex        =   0
      Top             =   3840
      Width           =   1095
   End
   Begin VB.Image Image1 
      Height          =   480
      Left            =   60
      Picture         =   "CutFiles.frx":0000
      Top             =   240
      Width           =   480
   End
End
Attribute VB_Name = "FMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const Title = "Cut File"
Const Extention = ".CFL"

Private Sub cboFileSize_Click()
  Select Case cboFileSize.ListIndex
    Case 0 '3½ floppy disk
      With txtFileSize
        .Text = 1457664
        .BackColor = BackColor
        .Locked = True
      End With
      
    Case 1 'custom
      With txtFileSize
        .BackColor = txtOutput.BackColor
        .Locked = False
      End With
  End Select
End Sub

Private Sub chkSendTo_Click()
  If chkSendTo.Value Then
    Me.txtAddress.Locked = False
    Me.txtAddress.BackColor = Me.txtSource.BackColor
  Else
    Me.txtAddress.Locked = True
    Me.txtAddress.BackColor = BackColor
  End If
End Sub

Private Sub cmdExecute_Click()
Dim hFile As Integer
Dim hOutput As Integer
Dim SourceFile As String
Dim OutputFile As String
Dim BatFile As String
Dim FileSize As Currency
Dim I As Integer
Dim J As Integer
Dim K As Integer
Dim Chunk As String
Dim OutputSize As Currency
Dim CopySize As Currency
Dim MaxSize As Currency

Dim BytesCopied As Long
Dim Progress As Single
Dim LastIndex As Integer

Dim MyBackColor As Long

  If Trim(txtSource) = "" Then
    MsgBox "Plese select the source file!", vbCritical, Title
    Exit Sub
  End If
  
  If Trim(Me.txtOutput) = "" Then
    MsgBox "Please decide the target file name!", vbCritical, Title
    Exit Sub
  End If
  
  cmdExecute.Enabled = False
  
  cmdFile.Enabled = False
  With txtSource
    MyBackColor = .BackColor
    .BackColor = BackColor
  End With
  
  Me.chkSendTo.Enabled = False
  
  With txtAddress
    .BackColor = BackColor
    .Locked = True
  End With
  
  With cboFileSize
    .Locked = True
    .BackColor = BackColor
  End With
  
  If cboFileSize.ListIndex <> 0 Then
    With txtFileSize
      .Locked = True
      .BackColor = BackColor
    End With
  End If
  
  SourceFile = Me.txtSource
  CopySize = 16384
  MaxSize = Me.txtFileSize
  
  FileSize = FileLen(SourceFile)
  hFile = FreeFile
  Open SourceFile For Binary As #hFile
  
  I = -1
  BytesCopied = 0
  Progress = 0
  LastIndex = 0
  K = 0
  While BytesCopied < FileSize
    If FileSize - BytesCopied < MaxSize Then
      MaxSize = FileSize - BytesCopied
    End If
    
    I = I + 1
    OutputSize = 0
    OutputFile = Me.txtOutput.Tag & _
                 Me.txtOutput.Text & _
                 "." & Format(I, "000") & _
                 Extention
    hOutput = FreeFile
    If Dir(OutputFile) <> "" Then
      Kill OutputFile
    End If
    Open OutputFile For Binary As #hOutput
    While OutputSize + CopySize <= MaxSize
      Chunk = Space(CopySize)
      Get #hFile, , Chunk
      Put #hOutput, , Chunk
      OutputSize = OutputSize + CopySize
      BytesCopied = BytesCopied + CopySize
      
      K = (K + 1) Mod 30
      Caption = "Cut Files " & String(K, "-")
      Progress = BytesCopied / FileSize * shpProgress.Count
      If LastIndex < Progress Then
        LastIndex = LastIndex + 1
        shpProgress(LastIndex - 1).Visible = True
      End If
      DoEvents
    Wend

    If (OutputSize <> MaxSize) Then
      Chunk = Space(MaxSize - OutputSize)
      Get #hFile, , Chunk
      Put #hOutput, , Chunk
      BytesCopied = BytesCopied + MaxSize - OutputSize
      
      J = (J + 1) Mod 30
      Caption = Title & " " & String(J, "-")
      Progress = BytesCopied / FileSize * shpProgress.Count
      If LastIndex < Progress Then
        LastIndex = LastIndex + 1
        shpProgress(LastIndex - 1).Visible = True
      End If
      DoEvents
    End If
    Close #hOutput
    
    If Me.chkSendTo.Value Then
      If Me.txtAddress <> "" Then
        SendEmail OutputFile, _
                  Me.txtOutput.Text & "." & Format(I, "000") & Extention, _
                  ""
      End If
    End If
  Wend
  Close #hFile
  
  '** creating bat file to combine all of the cut files
  hFile = FreeFile
  BatFile = Me.txtOutput.Tag & Me.txtOutput.Text & ".BAT"
  Open BatFile For Output As #hFile
  Print #hFile, "ECHO OFF"
  Print #hFile, "REN " & Me.txtOutput.Text & ".000" & Extention & " " & Me.txtOutput.Text
  If I >= 1 Then
    Print #hFile, "FOR %%A IN (" & Me.txtOutput.Text & ".*" & Extention & ") DO COPY " & _
                  Me.txtOutput.Text & " /B + %%A /B " & Me.txtOutput.Text & " /B"
    Print #hFile, "DEL " & Me.txtOutput.Text & ".*" & Extention
    Print #hFile, "DEL " & BatFile
  End If
  Close #hFile
  
  cmdExecute.Enabled = True
  cmdFile.Enabled = True
  Caption = Title
  With txtSource
    .BackColor = MyBackColor
  End With
  
  With txtOutput
    .BackColor = MyBackColor
    .Locked = False
  End With
  
  With cboFileSize
    .Locked = False
    .BackColor = MyBackColor
  End With
  
  Me.chkSendTo.Enabled = True
  If Me.chkSendTo.Value Then
    Me.txtAddress.Locked = False
    Me.txtAddress.BackColor = Me.txtOutput.BackColor
  End If
  
  If cboFileSize.ListIndex <> 0 Then
    With txtFileSize
      .Locked = False
      .BackColor = MyBackColor
    End With
  End If
  
  MsgBox "Cut file complete!", vbInformation, Title
  
  For J = 0 To shpProgress.Count - 1
    shpProgress(J).Visible = False
  Next J
End Sub

Private Sub cmdFile_Click()
Dim I As Integer

  With dlgFile
    .DialogTitle = "Input file ..."
    .CancelError = True
    .Filter = "Any file|*.*"
    .FilterIndex = 0
    On Error Resume Next
    .Action = 1
    If Err = 0 Then
      If .FileName <> "" Then
        Me.txtSource = .FileName
        Me.txtSource.ToolTipText = "File size: " & Format(FileLen(.FileName), "###,##0") & " bytes"
        Me.txtOutput = .FileTitle
        Me.txtOutput.Tag = Left(.FileName, Len(.FileName) - Len(.FileTitle))
      End If
    End If
  End With
  On Error GoTo 0
End Sub

Private Sub Form_Load()
  With cboFileSize
    .AddItem "3½ Floppy disk"
    .AddItem "Custom"
    .ListIndex = 0
  End With
End Sub

Private Sub SendEmail(ActualFile As String, FileName As String, Message As String, Optional Otherfile As String)
Dim oAppl As Outlook.Application
Dim oMail As Outlook.MailItem
Dim oTo As Outlook.Recipient
Dim oBCC As Outlook.Recipient

Dim I As Integer
Dim J As Integer
Dim strTemp As String

  Caption = "Send file to eMail ..."
  DoEvents
  
  Set oAppl = New Outlook.Application
  Set oMail = oAppl.CreateItem(olMailItem)
  With oMail
    Set oTo = .Recipients.Add(Trim(Me.txtAddress))
    .Subject = FileName
    .Attachments.Add ActualFile
    If Otherfile <> "" Then .Attachments.Add Otherfile
    .Body = Message
    .Send
    
    With tmrSendEmail
      .Tag = "SendEmail"
      .Interval = 5000 'five seconds
      .Enabled = True
    End With
  End With

  Set oMail = Nothing
  Set oAppl = Nothing
  
  '** delay five seconds
  Do
    DoEvents
  Loop Until tmrSendEmail.Tag = ""
End Sub

Private Sub tmrSendEmail_Timer()
  tmrSendEmail.Enabled = False
  tmrSendEmail.Tag = ""
End Sub
