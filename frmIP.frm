VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmIP 
   Caption         =   "Get your IP"
   ClientHeight    =   1950
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   2490
   LinkTopic       =   "Form1"
   ScaleHeight     =   1950
   ScaleWidth      =   2490
   StartUpPosition =   3  'Windows-Standard
   Begin MSWinsockLib.Winsock Winsock1 
      Left            =   1200
      Top             =   600
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   120
      TabIndex        =   3
      Top             =   360
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Show me - I don't belive it..."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Line Line1 
      BorderColor     =   &H8000000C&
      X1              =   120
      X2              =   2400
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Your external IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   840
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Your internal IP:"
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmIP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
    Text1.Text = GetInternetIP(False)
    Text2.Text = GetInternetIP(True)
End Sub


