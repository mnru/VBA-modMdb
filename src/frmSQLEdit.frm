VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSQLEdit 
   Caption         =   "exec SQL"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   4560
   OleObjectBlob   =   "frmSQLEdit.frx":0000
   StartUpPosition =   1  'オーナー フォームの中央
End
Attribute VB_Name = "frmSQLEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdDisplay_Click()
    Call displayQueryTable(, Me.tbxSQL.Value, formDBPath)
End Sub

Private Sub cmdExec_Click()
    Call execSQL(Me.tbxSQL.Value, formDBPath)
    MsgBox "finished"
End Sub

Private Sub cmdHide_Click()
    Call Me.Hide
End Sub
