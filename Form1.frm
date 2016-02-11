VERSION 5.00
Object = "{20C62CAE-15DA-101B-B9A8-444553540000}#1.1#0"; "MSMAPI32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   465
      Left            =   1800
      TabIndex        =   0
      Top             =   2160
      Width           =   1335
   End
   Begin MSMAPI.MAPIMessages MAPIMessages1 
      Left            =   810
      Top             =   1470
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      AddressEditFieldCount=   1
      AddressModifiable=   0   'False
      AddressResolveUI=   0   'False
      FetchSorted     =   0   'False
      FetchUnreadOnly =   0   'False
   End
   Begin MSMAPI.MAPISession MAPISession1 
      Left            =   690
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      DownloadMail    =   -1  'True
      LogonUI         =   -1  'True
      NewSession      =   0   'False
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Dim asAddr() As String
    Dim i As Integer
        
    Command1.Enabled = False
    
    'fill vector to send to multiple addresses
    asAddr = Split("Bill Gates <bill@microsoft.com>|Barack Obama <bobama@usa.net>|Martin <martin@habitech.com.ar>", "|")
    
    MAPISession1.SignOn
    MAPIMessages1.SessionID = MAPISession1.SessionID
    MAPIMessages1.Compose
    For i = 0 To UBound(asAddr)
        MAPIMessages1.RecipIndex = i
        MAPIMessages1.RecipDisplayName = asAddr(i)
        MAPIMessages1.ResolveName
        Debug.Print "Nombre: " & MAPIMessages1.RecipDisplayName & ", email: " & MAPIMessages1.RecipAddress
    Next i
    MAPIMessages1.MsgSubject = "Subject " & App.Title
    MAPIMessages1.MsgNoteText = "<html><body><h3><font color=grey>html</font>&nbsp<font color=blue>mail</font></h1></body></html>"
    
    'MAPIMessages1.AttachmentIndex = 0
    'MAPIMessages1.AttachmentPathName = "c:\somefilename.txt"
    'MAPIMessages1.AttachmentIndex = 1
    'MAPIMessages1.AttachmentPathName = "c:\otherfilename.txt"
    
    'vdialog = false : sends the email silently / vdialog = true: activates webmail client browser window , to search for the email in the drafts folder.
    MAPIMessages1.Send False
    MAPISession1.SignOff
        
    Command1.Enabled = True
End Sub
