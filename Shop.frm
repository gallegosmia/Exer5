VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form Form2 
   BackColor       =   &H008080FF&
   Caption         =   "Form2"
   ClientHeight    =   9105
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   9345
   LinkTopic       =   "Form2"
   ScaleHeight     =   9105
   ScaleWidth      =   9345
   StartUpPosition =   3  'Windows Default
   Begin VB.OptionButton Optionno 
      Caption         =   "NO"
      Height          =   375
      Left            =   3840
      TabIndex        =   18
      Top             =   2520
      Width           =   975
   End
   Begin VB.TextBox txtprice 
      Height          =   495
      Left            =   2400
      TabIndex        =   15
      Top             =   1920
      Width           =   4095
   End
   Begin VB.TextBox txtcategory 
      Height          =   495
      Left            =   2400
      TabIndex        =   14
      Top             =   1320
      Width           =   4095
   End
   Begin VB.CommandButton btnClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   5640
      TabIndex        =   10
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton btnDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4080
      TabIndex        =   9
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton btnEdit 
      Caption         =   "Edit"
      Height          =   495
      Left            =   2520
      TabIndex        =   8
      Top             =   8280
      Width           =   1215
   End
   Begin VB.CommandButton btnAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   960
      TabIndex        =   7
      Top             =   8280
      Width           =   1215
   End
   Begin MSDataGridLib.DataGrid DataGrid1 
      Height          =   3255
      Left            =   360
      TabIndex        =   6
      Top             =   4920
      Width           =   7335
      _ExtentX        =   12938
      _ExtentY        =   5741
      _Version        =   393216
      HeadLines       =   1
      RowHeight       =   15
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   13321
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtsearch 
      Height          =   495
      Left            =   1680
      TabIndex        =   5
      Top             =   4320
      Width           =   4335
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Product Details"
      BeginProperty Font 
         Name            =   "Papyrus"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3735
      Left            =   720
      TabIndex        =   0
      Top             =   240
      Width           =   6495
      Begin VB.OptionButton Optionyes 
         Caption         =   "YES"
         Height          =   375
         Left            =   1680
         TabIndex        =   17
         Top             =   2280
         Width           =   975
      End
      Begin VB.TextBox txtbrand 
         Height          =   495
         Left            =   1680
         TabIndex        =   13
         Top             =   480
         Width           =   4095
      End
      Begin MSComCtl2.DTPicker dtexpiration 
         Height          =   495
         Left            =   1680
         TabIndex        =   12
         Top             =   2880
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         _Version        =   393216
         Format          =   114556929
         CurrentDate     =   45934
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Vegan"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   16
         Top             =   2280
         Width           =   1215
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Expiration Date"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   11
         Top             =   2880
         Width           =   1215
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Price"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   3
         Top             =   1680
         Width           =   1215
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Category"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   2
         Top             =   1080
         Width           =   1215
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Brand Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   240
         TabIndex        =   1
         Top             =   480
         Width           =   1215
      End
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00FFFFFF&
      Caption         =   "Search Here"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   360
      TabIndex        =   4
      Top             =   4320
      Width           =   1215
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub btnAdd_Click()

    On Error GoTo btnAdd_Click_Err

100     If btnAdd.Caption = "Add" Then
102         btnAdd.Caption = "Save"
104         btnClose.Caption = "Cancel"
108         txtbrand.Enabled = True
            txtcategory.Enabled = True
112         txtsearch.Enabled = True
            txtprice.Enabled = True
            dtexpiration.Enabled = True
            Optionyes.Enabled = True
            Optionno.Enabled = True
120         DataGrid1.Enabled = False
        Else
             If txtbrand.Text = "" Or _
                txtcategory.Text = "" Or _
                dtexpiration = "" Or _
                (Optionyes.Value = False And Optionno.Value = False) Or _
                txtprice.Text = "" Then
124             MsgBox _
                    "All fields are required!", _
                    vbInformation, _
                    "Webplus Lending Corporation"
            Else
138             If MsgBox( _
                        "Are you sure you want to add new Record?", _
                        vbQuestion + vbYesNo, _
                        "Webplus Lending Corporation") _
                        = vbYes Then
                    Call MakeupProd
140                 With rsMakeupProd
142                     .AddNew
146                     !BrandName = Trim$( _
                                txtbrand.Text)
148                     !Category = Trim$( _
                                txtcategory.Text)
                            !ExpirationDate = dtexpiration
                            !Vegan = Optionyes
                            !Vegan = Optionno
                            !Price = Trim$( _
                                Val(txtprice.Text))
160                     .Update
                    End With
182                 MsgBox _
                            "Record Successfully Added", _
                            vbInformation, _
                            "Webplus Lending Corporation"
184                 Unload Me
                    Me.Show
                End If
            End If
        End If

    Exit Sub

btnAdd_Click_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.btnAdd_Click", _
        Erl
    Resume Next

End Sub

Private Sub btnClose_Click()

    On Error GoTo btnClose_Click_Err

100     If btnClose.Caption = "Close" Then
102         Unload Me
        Else
104         btnClose.Caption = "Close"
106         txtbrand.Enabled = True
            txtcategory.Enabled = True
112         txtsearch.Enabled = True
            txtprice.Enabled = True
            dtexpiration.Enabled = True
            Optionyes.Enabled = True
            Optionno.Enabled = True
120         DataGrid1.Enabled = False
        End If

    Exit Sub

btnClose_Click_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.btnClose_Click", _
        Erl
    Resume Next

End Sub

Private Sub btnDelete_Click()

    On Error GoTo btnDelete_Click_Err

116     If MsgBox( _
                "Are you sure you want to delete this user?", _
                vbQuestion + vbYesNo) = vbYes _
                Then
118         rsMakeupProd.Delete
120         rsMakeupProd.Update
122         MsgBox _
                    "Record successfully deleted", _
                    vbInformation
124         Unload Me
126         Me.Show
        End If

    Exit Sub

btnDelete_Click_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.btnDelete_Click", _
        Erl
    Resume Next

End Sub

Private Sub btnEdit_Click()

100     If btnEdit.Caption = "Edit" Then
102         btnEdit.Caption = "Update"
104         txtbrand.Enabled = True
106         txtcategory.Enabled = True
108         txtsearch.Enabled = True
110         dtexpiration.Enabled = True
            Optionyes.Enabled = True
            Optionno.Enabled = True
116         DataGrid1.Enabled = False
        Else
124         If txtbrand.Text = "" Or _
                txtcategory.Text = "" Or _
                (Optionyes.Value = False And Optionno.Value = False) Or _
                dtexpiration = "" Then
126             MsgBox _
                    "All fiends are required! ", _
                    vbInformation
            Else
128             If MsgBox( _
                    "Are you sure to update this record?", _
                    vbQuestion + vbYesNo, _
                    "J Lending Corporation") = _
                    vbYes Then
130                 With rsMakeupProd
136                     !BrandName = Trim$( _
                                txtbrand.Text)
148                     !Category = Trim$( _
                                txtcategory.Text)
                        !ExpirationDate = dtexpiration
                        !Vegan = Optionno
                        !Vegan = Optionyes
                        !Price = Trim$( _
                            Val(txtprice.Text))
                        .Update
                    End With
                    btnEdit.Caption = "Edit"
                    rsMakeupProd.Close
                End If
                Unload Me
                Me.Show
            End If
        End If

End Sub

Private Sub DataGrid1_Click()

    On Error GoTo DataGrid1_Click_Err

100     If rsMakeupProd.RecordCount = 0 Then
        Else
102         btnAdd.Enabled = False
104         btnClose.Caption = "Cancel"
106         btnEdit.Enabled = True
108         btnDelete.Enabled = True

110         With rsMakeupProd
120             txtbrand.Text = !BrandName
122             txtcategory.Text = !Category
123             txtprice.Text = !Price
124             Optionyes.Value = !Vegan
125             Optionno.Value = !Vegan
            If IsNull(!ExpirationDate) Then
            Else
126             dtexpiration.Value = !ExpirationDate
            End If
            End With
        End If

    Exit Sub

DataGrid1_Click_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.DataGrid1_Click", _
        Erl
    Resume Next

End Sub

Private Sub DataGrid1_RowColChange(LastRow As Variant, ByVal LastCol As Integer)
100    Call DataGrid1_Click
End Sub

Private Sub Form_Load()

    On Error GoTo Form_Load_Err

108     txtbrand.Enabled = True
        txtcategory.Enabled = True
114     txtsearch.Enabled = True
        dtexpiration.Enabled = True
        Optionyes.Enabled = True
        Optionno.Enabled = True
120     DataGrid1.Enabled = True

    Call connect
    Call MakeupProd

    Set DataGrid1.DataSource = rsMakeupProd
130     DataGrid1.Width = 19000

    Exit Sub

Form_Load_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.Form_Load", _
        Erl
    Resume Next

End Sub

Private Sub Optionno_Click()
    Optionno.Value = True
    Optionyes.Value = False
End Sub

Private Sub Optionyes_Click()
    Optionyes.Value = True
    Optionno.Value = False
End Sub

Private Sub txtsearch_Change()

    Set DataGrid1.DataSource = rsMakeupProd
    Exit Sub

txtsearch_Change_Err:
    ErrReport Err.Description, _
        "LendingClientV2.frm_Users.txtsearch_Change", _
        Erl
    Resume Next

    '</EhFooter>

End Sub


