VERSION 5.00
Begin VB.Form Form1 
   BackColor       =   &H008080FF&
   Caption         =   "Form1"
   ClientHeight    =   3930
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8925
   LinkTopic       =   "Form1"
   ScaleHeight     =   3930
   ScaleWidth      =   8925
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton btnSeller 
      Caption         =   "SELLER"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   4560
      TabIndex        =   2
      Top             =   1560
      Width           =   2775
   End
   Begin VB.CommandButton btnCustomer 
      Caption         =   "CUSTOMER"
      BeginProperty Font 
         Name            =   "Oswald"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1335
      Left            =   1320
      TabIndex        =   1
      Top             =   1560
      Width           =   2775
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H00C0C0FF&
      Caption         =   "AIM HIGH MAKE UP STORE"
      BeginProperty Font 
         Name            =   "Perpetua Titling MT"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   975
      Left            =   2535
      TabIndex        =   0
      Top             =   240
      Width           =   3855
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Label2_Click()

End Sub

Private Sub Command2_Click()

End Sub

Private Sub btnCustomer_Click()
Unload Me
End Sub

Private Sub btnSeller_Click()
Form2.Show
Unload Me
End Sub
