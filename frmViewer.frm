VERSION 5.00
Begin VB.Form frmViewer 
   Caption         =   "Hex View"
   ClientHeight    =   8760
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   10245
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmViewer.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8760
   ScaleWidth      =   10245
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdSaveToFile 
      Caption         =   "Save To File"
      Height          =   495
      Left            =   2220
      TabIndex        =   2
      Top             =   120
      Width           =   1995
   End
   Begin VB.CommandButton cmdCopyAll 
      Caption         =   "Copy All"
      Height          =   495
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   1995
   End
   Begin VB.TextBox txtHex 
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   7875
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   660
      Width           =   9975
   End
End
Attribute VB_Name = "frmViewer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Public Sub ShowHex(ByVal sTitle As String, ByRef bData() As Byte)
    Me.Caption = "Hex Inspector - " & sTitle
    ' Show up to 64 KB in the textbox
    txtHex.Text = FormatHexDump(bData, 65536)
    Me.Show vbModal, frmMain
End Sub

Private Sub cmdCopyAll_Click()
    Clipboard.Clear
    Clipboard.SetText txtHex
    MsgBox "Hex Data Copied to Clipboard!", vbInformation
End Sub

Private Sub cmdSaveToFile_Click()
    Dim S As String

    S = ShowSave("Text Files|*.txt|All Files|*.*", "", "Select File To Save Hex Data To", "txt")
    If Len(S) = 0 Then Exit Sub

    If FileExist(S) Then
        If MsgBox("Are you sure you want to overwrite this file - " & FileGetName(S), vbQuestion + vbYesNo) = vbNo Then
            Exit Sub
        End If
    End If

    If FileSave(txtHex, S) Then
        MsgBox "Hex data successfully saved to " & FileGetName(S), vbInformation
    Else
        MsgBox "There was a problem saving the data to the file " & FileGetName(S), vbCritical
    End If

End Sub

Private Sub Form_Load()
   Set Icon = frmMain.Icon
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    txtHex.Move 25, 800, Me.ScaleWidth - 100, Me.ScaleHeight - 900
End Sub
