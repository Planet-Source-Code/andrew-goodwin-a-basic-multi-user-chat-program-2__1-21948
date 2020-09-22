VERSION 5.00
Begin VB.Form frmConnect 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Basic Multi User Chat Program"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4080
   Icon            =   "frmConnect.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   4080
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton txtCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   2640
      TabIndex        =   3
      Top             =   960
      Width           =   1095
   End
   Begin VB.CommandButton txtConnect 
      Caption         =   "Connect"
      Height          =   375
      Left            =   1440
      TabIndex        =   2
      Top             =   960
      Width           =   1095
   End
   Begin VB.TextBox txtIPAddress 
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   3495
   End
   Begin VB.Label Label1 
      Caption         =   "Enter IP Address Of Server:"
      Height          =   255
      Left            =   240
      TabIndex        =   1
      Top             =   240
      Width           =   3495
   End
End
Attribute VB_Name = "frmConnect"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub txtCancel_Click()

'Hide the form from view
frmConnect.Visible = False

End Sub

Private Sub txtConnect_Click()

'Set the remote port, this has to be the same as the server
'port
frmMultiChat.wsClient.RemotePort = 789

'Connect to the given IP address
frmMultiChat.wsClient.Connect txtIPAddress.Text
DoEvents

'**This is all that is needed to connect the client to the server
'**the rest below is just handeling menu buttons and other stuff**

'**********************************************************************
'**********************************************************************

'Hide the from now that it has been used
frmConnect.Visible = False

'Set the forms caption so that we know what we are
'connected as
frmMultiChat.Caption = "Connected As Client"

'As seen as we are the client we cant connect as the server
'so disable the Start Server button
frmMultiChat.mnuConnectAsClient.Enabled = False

'We have connected as the client so disable the Connect As'
'Client button
frmMultiChat.mnuStartServer.Enabled = False

'We have started connected to the server so enable the
'End Connection Button
frmMultiChat.mnuEndConnection.Enabled = True

End Sub
