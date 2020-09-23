VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form fPEinfO 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "[ PEinfO ] by Sangaletti Federico"
   ClientHeight    =   5655
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7455
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5655
   ScaleWidth      =   7455
   StartUpPosition =   2  'CenterScreen
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   6600
      Top             =   840
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   16777215
      ImageWidth      =   72
      ImageHeight     =   72
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3D12
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":728C
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":AF9E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog CommonDialog1 
      Left            =   3240
      Top             =   480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3840
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":ECB0
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":10802
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":12354
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":13EA6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1050
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   7455
      _ExtentX        =   13150
      _ExtentY        =   1852
      ButtonWidth     =   2011
      ButtonHeight    =   1799
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   4
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Analyze PE"
            Key             =   "kLoad"
            Object.ToolTipText     =   "Load and analyze PE"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hdrs structure"
            Object.ToolTipText     =   "Headers structure"
            ImageIndex      =   2
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   7
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kDosHeader"
                  Text            =   "IMAGE_DOS_HEADER"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kNTHeaders"
                  Text            =   "IMAGE_NT_HEADERS"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kFileHeader"
                  Text            =   "IMAGE_FILE_HEADER"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kOptionalHeader"
                  Text            =   "IMAGE_OPTIONAL_HEADER"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kSectionHeader"
                  Text            =   "IMAGE_SECTION_HEADER"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kExportDir"
                  Text            =   "IMAGE_EXPORT_DIRECTORY"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kImportDescr"
                  Text            =   "IMAGE_IMPORT_DECRIPTOR"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Tools"
            Object.ToolTipText     =   "Tools"
            ImageIndex      =   3
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   4
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kRvaToOffset"
                  Text            =   "RVA to Offset"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kAlignedSize"
                  Text            =   "Aligned size"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kDec2Hex"
                  Text            =   "Dec to Hex"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "kHex2Dec"
                  Text            =   "Hex to Dec"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Key             =   "kExit"
            ImageIndex      =   4
         EndProperty
      EndProperty
      BorderStyle     =   1
   End
   Begin MSComctlLib.ListView lstReport 
      Height          =   4335
      Left            =   120
      TabIndex        =   1
      Top             =   1200
      Width           =   7215
      _ExtentX        =   12726
      _ExtentY        =   7646
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      PictureAlignment=   3
      _Version        =   393217
      Icons           =   "ImageList2"
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
      Picture         =   "Form1.frx":159F8
   End
End
Attribute VB_Name = "fPEinfO"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'##########################################################
'##########################################################
'######## Title:. PEInfO [26/09/05]                ########
'######## Author: Sangaletti Federico              ########
'######## e-mail: sangaletti@aliceposta.it         ########
'########------------------------------------------########
'########   !IF YOU LIKE THIS CODE PLEASE VOTE!    ########
'########------------------------------------------########
'##########################################################
'##########################################################


Private Sub Form_Load()
    Load fShowInfo
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Unload fShowInfo
End Sub

Private Sub lstReport_DblClick()
    On Error Resume Next
    Select Case lstReport.SelectedItem.Key
        Case "kDOSHeader":
            fShowInfo.Caption = "IMAGE_DOS_HEADER"
            fShowInfo.txtInfo = DOS_HEADER_INFO
            fShowInfo.Show
        
        Case "kNTHeaders":
            fShowInfo.Caption = "IMAGE_NT_HEADERS"
            fShowInfo.txtInfo = NT_HEADERS_INFO
            fShowInfo.Show
            
        Case "kSectionHeader":
            fShowInfo.Caption = "IMAGE_SECTION_HEADER"
            fShowInfo.txtInfo = SECTION_TABLE
            fShowInfo.Show
            
        Case "kExportTable":
            fShowInfo.Caption = "IMAGE_EXPORT_DIRECTORY"
            fShowInfo.txtInfo = EXPORT_TABLE
            fShowInfo.Show
            
        Case "kImportTable"
            fShowInfo.Caption = "IMAGE_IMPORT_DESCRIPTOR"
            fShowInfo.txtInfo = IMPORT_TABLE
            fShowInfo.Show
    End Select
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "kLoad":
            CommonDialog1.FileName = vbNullString
            CommonDialog1.Filter = "PE [Portable Executable]|*.exe;*.dll;*.ocx"
            CommonDialog1.ShowOpen
            If CommonDialog1.FileName <> vbNullString Then GetPEInfo CommonDialog1.FileName
        
        Case "kExit":
            MsgBox "If you like this code please vote for it.", vbInformation
            End
    End Select
End Sub

Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
        Case "kDosHeader":
            fShowInfo.Caption = "IMAGE_DOS_HEADER Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_DOS_HEADER
            fShowInfo.Show
            
        Case "kNTHeaders":
            fShowInfo.Caption = "IMAGE_NT_HEADERS Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_NT_HEADERS
            fShowInfo.Show
        
        Case "kFileHeader":
            fShowInfo.Caption = "IMAGE_FILE_HEADER Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_FILE_HEADER
            fShowInfo.Show
            
        Case "kOptionalHeader":
            fShowInfo.Caption = "IMAGE_OPTIONAL_HEADER Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_OPTIONAL_HEADER
            fShowInfo.Show
            
        Case "kSectionHeader":
            fShowInfo.Caption = "IMAGE_SECTION_HEADER Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_SECTION_HEADER
            fShowInfo.Show
            
        Case "kExportDir":
            fShowInfo.Caption = "IMAGE_EXPORT_DIRECTORY Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_EXPORT_DIRECTORY
            fShowInfo.Show
            
        Case "kImportDescr":
            fShowInfo.Caption = "IMAGE_IMPORT_DESCRIPTOR Structure (C syntax)"
            fShowInfo.txtInfo = TXT_IMAGE_IMPORT_DESCRIPTOR
            fShowInfo.Show
            
            
            
        
        Case "kRvaToOffset":
            If SECTION_TABLE <> "" Then
                MsgBox "Offset is 0x" & Hex$(RVAToOffset(SectionHeaders, CLng("&H" & InputBox("Type a RVA (Relative Virtual Address) in hexadecimal.")))), vbInformation
            Else
                MsgBox "Open and analyze a PE first!", vbExclamation
            End If
        
        Case "kAlignedSize":
            MsgBox "Aligned size is 0x" & Hex$(GetAlignedSize(CLng("&H" & InputBox("Type real size in hex value.")), CLng("&H" & InputBox("Type alignment in hex value")))), vbInformation
            
        Case "kDec2Hex":
            MsgBox "Hexadecimal value is 0x" & Hex$(Val(InputBox("Type a decimal value."))), vbInformation
            
        Case "kHex2Dec":
            MsgBox "Decimal value is " & CLng("&H" & InputBox("Type an hexadecimal value.")), vbInformation
    End Select
End Sub
