VERSION 5.00
Begin VB.Form frmExportOpts 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Export Options"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6615
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmExportOpts.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3360
   ScaleWidth      =   6615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdOK 
      Caption         =   "Ok"
      Height          =   555
      Left            =   4620
      TabIndex        =   5
      Top             =   2460
      Width           =   1635
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   555
      Left            =   2880
      TabIndex        =   4
      Top             =   2460
      Width           =   1635
   End
   Begin VB.CheckBox chkSkipBlobs 
      Caption         =   "Skip raw unparsed binary (.bin) blobs during Export All"
      Height          =   315
      Left            =   240
      TabIndex        =   3
      Top             =   1680
      Value           =   1  'Checked
      Width           =   6135
   End
   Begin VB.CheckBox chkGroupFolders 
      Caption         =   "Group exported files into subfolders by Form / Designer name"
      Height          =   315
      Left            =   240
      TabIndex        =   2
      Top             =   1200
      Width           =   6135
   End
   Begin VB.CheckBox chkImageStrip 
      Caption         =   "Export ImageLists / multi-image entries as a horizontal PNG strip"
      Height          =   315
      Left            =   240
      TabIndex        =   1
      Top             =   720
      Width           =   6135
   End
   Begin VB.CheckBox chkPackIcons 
      Caption         =   "Combine multi-resolution icons into a single .ico file"
      Height          =   315
      Left            =   240
      TabIndex        =   0
      Top             =   300
      Width           =   4935
   End
End
Attribute VB_Name = "frmExportOpts"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mConfirmed As Boolean

Public Function PromptExportOptions(ByVal frmOwner As Form, Optional ByVal bIsBatchAll As Boolean = True) As Boolean
    mConfirmed = False
    
    ' Load saved defaults (using your GetStr helper)
    chkPackIcons.Value = IIf(GetInt("OptPackIcons", 1) = 1, vbChecked, vbUnchecked)
    chkImageStrip.Value = IIf(GetInt("OptImageStrip", 0) = 1, vbChecked, vbUnchecked)
    chkGroupFolders.Value = IIf(GetInt("OptGroupFolders", 1) = 1, vbChecked, vbUnchecked)
    chkSkipBlobs.Value = IIf(GetInt("OptSkipBlobs", 1) = 1, vbChecked, vbUnchecked)
    
    ' Disable batch-only folder/blob options when exporting a single file
    chkGroupFolders.Enabled = bIsBatchAll
    chkSkipBlobs.Enabled = bIsBatchAll
    
    Me.Show vbModal, frmOwner
    PromptExportOptions = mConfirmed
End Function

Private Sub cmdOK_Click()
    ' Save to global UDT
    With gExportOpts
        .PackMultiIcons = (chkPackIcons.Value = vbChecked)
        .ExportImageListsAsStrip = (chkImageStrip.Value = vbChecked)
        .GroupByFormFolder = (chkGroupFolders.Value = vbChecked)
        .SkipUnindexedBlobs = (chkSkipBlobs.Value = vbChecked)
    End With
    
    ' Persist preferences for next session
    SaveInt "OptPackIcons", IIf(gExportOpts.PackMultiIcons, 1, 0)
    SaveInt "OptImageStrip", IIf(gExportOpts.ExportImageListsAsStrip, 1, 0)
    SaveInt "OptGroupFolders", IIf(gExportOpts.GroupByFormFolder, 1, 0)
    SaveInt "OptSkipBlobs", IIf(gExportOpts.SkipUnindexedBlobs, 1, 0)
    
    gExportOpts.OptionsCancelled = False
    mConfirmed = True
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    mConfirmed = False
    gExportOpts.OptionsCancelled = True
    Unload Me
End Sub

Private Sub Form_Load()
   Set Icon = frmMain.Icon
End Sub
