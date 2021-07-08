VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form ClientFRM 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Adam's Chat   [Client]"
   ClientHeight    =   4200
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7050
   LinkTopic       =   "Form3"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4200
   ScaleWidth      =   7050
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox MainTXT 
      Height          =   3615
      Left            =   120
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      Top             =   120
      Width           =   6855
   End
   Begin VB.TextBox dataTXT 
      Height          =   285
      Left            =   120
      MaxLength       =   100
      TabIndex        =   1
      Top             =   3840
      Width           =   5655
   End
   Begin VB.CommandButton senddataCMD 
      Caption         =   "&Send"
      Default         =   -1  'True
      Height          =   315
      Left            =   5880
      TabIndex        =   0
      Top             =   3840
      Width           =   975
   End
   Begin MSWinsockLib.Winsock Winsock 
      Left            =   600
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "© 2000 One"
      BeginProperty Font 
         Name            =   "MS Serif"
         Size            =   6
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   135
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   855
   End
   Begin VB.Label nickCLIENT 
      Height          =   135
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label nickSERVER 
      Height          =   135
      Left            =   120
      TabIndex        =   3
      Top             =   4320
      Visible         =   0   'False
      Width           =   255
   End
End
Attribute VB_Name = "ClientFRM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'===========================================================================
'======================Displays info if disconnected========================
Private Sub Winsock_Close()
    MainTXT.SelText = "¨'°*·º·*°'¨ Disconnected to Server ¨'°*·º·*°'¨" & vbCrLf 'displays text if u get disconnected from the server
    senddataCMD.Enabled = False 'this wont let you send text anymore since server is disconnected
End Sub

'===========================================================================
'=======================RETREIVES DATA FROM SERVER==========================
Private Sub winsock_DataArrival(ByVal bytesTotal As Long)
Dim strData, strData2 As String 'where the data sent by the client will be stored
Call Winsock.GetData(strData, vbString) 'gets the data sent by the client

strData2 = Left(strData, 1) 'saves the first variable's value to strData2
strData = Mid(strData, 2) 'saves the text the server sent to strData

If strData2 = "C" Then 'tells form to load and do other form stuff
    nickSERVER.Caption = strData 'loads the server's username from data sent
    Me.Show 'shows Client frm
    Unload ConnectFRM 'unloads the connect frm
    Winsock.SendData "N" & nickCLIENT.Caption 'sends your username to the server
    ClientFRM.Caption = "Adam's Chat   [Welcome, " & nickCLIENT.Caption & "!]" 'renames the form approiately
    MainTXT.SelText = "¨'°*·º·*°'¨ Connected to Server ¨'°*·º·*°'¨" & vbCrLf 'displays that connection worked
End If

If strData2 = "T" Then MainTXT.SelText = nickSERVER.Caption & ":     " & strData & vbCrLf 'adds the data to the txtbox from the server

End Sub
'===========================================================================
'========================SENDS DATA TYPED TO SERVER=========================
Private Sub senddataCMD_Click()
MainTXT.SelText = nickCLIENT.Caption & ":     " & dataTXT.Text & vbCrLf 'puts what u typed in ur maintxt
Winsock.SendData "T" & dataTXT.Text 'sends the data to the server
dataTXT.Text = "" 'clears the txtbox u typed in

End Sub
'===========================================================================
