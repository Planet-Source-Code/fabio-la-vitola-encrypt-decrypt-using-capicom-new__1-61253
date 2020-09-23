VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Encrypt and Decrypt using Microsoft™ CAPICOM 2.1"
   ClientHeight    =   4155
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5550
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4155
   ScaleWidth      =   5550
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin TabDlg.SSTab SSTab1 
      Height          =   4065
      Left            =   60
      TabIndex        =   0
      Top             =   45
      Width           =   5445
      _ExtentX        =   9604
      _ExtentY        =   7170
      _Version        =   393216
      Tabs            =   2
      Tab             =   1
      TabsPerRow      =   2
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Encrypt Text"
      TabPicture(0)   =   "frmMain.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "Picture1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "Encrypt File"
      TabPicture(1)   =   "frmMain.frx":0028
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "comDatabase"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "Picture2"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      Begin VB.PictureBox Picture2 
         Height          =   3600
         Left            =   75
         ScaleHeight     =   3540
         ScaleWidth      =   5235
         TabIndex        =   15
         Top             =   375
         Width           =   5295
         Begin VB.CommandButton cmdFileEncrypt 
            Caption         =   "File Encrypt"
            Height          =   315
            Left            =   855
            TabIndex        =   20
            ToolTipText     =   "Encrypt the file !"
            Top             =   795
            Width           =   1350
         End
         Begin VB.TextBox Text1 
            Height          =   285
            Left            =   870
            TabIndex        =   19
            ToolTipText     =   "The password neede to perform everything"
            Top             =   390
            Width           =   2400
         End
         Begin VB.CommandButton cmdFileDecrypt 
            Caption         =   "File Decrypt"
            Height          =   315
            Left            =   2310
            TabIndex        =   18
            ToolTipText     =   "Decrypt the file !"
            Top             =   795
            Width           =   1350
         End
         Begin VB.TextBox Text2 
            Height          =   285
            Left            =   870
            TabIndex        =   17
            ToolTipText     =   "Here the file name you will encrypt or decrypt"
            Top             =   45
            Width           =   3930
         End
         Begin VB.CommandButton Command3 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4860
            Picture         =   "frmMain.frx":0044
            Style           =   1  'Graphical
            TabIndex        =   16
            ToolTipText     =   "Select the file to be Encrypted/decrypted"
            Top             =   45
            Width           =   315
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Left            =   60
            TabIndex        =   22
            Top             =   450
            Width           =   690
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "File Name"
            Height          =   195
            Left            =   75
            TabIndex        =   21
            Top             =   75
            Width           =   690
         End
      End
      Begin VB.PictureBox Picture1 
         Height          =   3630
         Left            =   -74940
         ScaleHeight     =   3570
         ScaleWidth      =   5250
         TabIndex        =   1
         Top             =   375
         Width           =   5310
         Begin VB.Frame Frame1 
            BeginProperty Font 
               Name            =   "Tahoma"
               Size            =   1.5
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   45
            Left            =   -165
            TabIndex        =   14
            Top             =   1770
            Width           =   5325
         End
         Begin VB.TextBox txtOriginal 
            Height          =   285
            Left            =   1320
            TabIndex        =   9
            ToolTipText     =   "This is the Decrypted text"
            Top             =   3195
            Width           =   3885
         End
         Begin VB.TextBox txtBody 
            Height          =   870
            Left            =   105
            MultiLine       =   -1  'True
            TabIndex        =   8
            ToolTipText     =   "Insert here the text you want to decrypt"
            Top             =   1920
            Width           =   5130
         End
         Begin VB.TextBox txtPwd 
            Height          =   285
            Left            =   1335
            TabIndex        =   7
            ToolTipText     =   "Decrypting password"
            Top             =   2880
            Width           =   2955
         End
         Begin VB.CommandButton cmdDecrypt 
            Caption         =   "Decrypt"
            Height          =   315
            Left            =   4380
            TabIndex        =   6
            ToolTipText     =   "Decrypts the text !"
            Top             =   2895
            Width           =   840
         End
         Begin VB.CommandButton cmdEncrypt 
            Caption         =   "Encrypt"
            Height          =   315
            Left            =   4395
            TabIndex        =   5
            ToolTipText     =   "Encrypts the text !"
            Top             =   390
            Width           =   840
         End
         Begin VB.TextBox txtEncrypted 
            Height          =   945
            Left            =   105
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   4
            ToolTipText     =   "This is the Encrypted text"
            Top             =   750
            Width           =   5130
         End
         Begin VB.TextBox txtPassword 
            Height          =   285
            Left            =   1365
            TabIndex        =   3
            ToolTipText     =   "Encrypting password"
            Top             =   390
            Width           =   2955
         End
         Begin VB.TextBox txtMessage 
            Height          =   285
            Left            =   1365
            TabIndex        =   2
            ToolTipText     =   "Insert here the text you want to encrypt"
            Top             =   90
            Width           =   3885
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "Text"
            Height          =   195
            Left            =   120
            TabIndex        =   13
            Top             =   3270
            Width           =   330
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Left            =   120
            TabIndex        =   12
            Top             =   2940
            Width           =   690
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Password"
            Height          =   195
            Left            =   120
            TabIndex        =   11
            Top             =   435
            Width           =   690
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "Text to encrypt"
            Height          =   195
            Left            =   120
            TabIndex        =   10
            Top             =   90
            Width           =   1125
         End
      End
      Begin MSComDlg.CommonDialog comDatabase 
         Left            =   3015
         Top             =   2595
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         Flags           =   4
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim TheCodec As New Codifier

Private Sub cmdEncrypt_Click()
txtEncrypted = TheCodec.EncryptText(txtMessage, txtPassword)            'Text Encryption using new object
End Sub

Private Sub cmdDecrypt_Click()
txtOriginal = TheCodec.DecryptText(txtBody, txtPwd)                     'Text Decryption using new object
End Sub

Private Sub cmdFileDecrypt_Click()
TheCodec.DecryptFile Text2, Text2 & ".dec", Text1                       'File Decryption using new object
End Sub

Private Sub cmdFileEncrypt_Click()
TheCodec.EncryptFile Text2, Text2 & ".enc", Text1                       'File Encryption using new object
End Sub

Private Sub Command3_Click()
On Error Resume Next                                                    'Just to prevent boring "Cancel" errors or similar

comDatabase.FileName = Text2                                            'Save the selected file name for further selections
comDatabase.ShowOpen                                                    'Re-open file list
If comDatabase.FileName <> vbNullString Then
   Text2 = comDatabase.FileName                                         'if a valid file name is provided
   'You can also enable or disable button if u want !
 Else
   'Do whatever u want !
End If
End Sub

