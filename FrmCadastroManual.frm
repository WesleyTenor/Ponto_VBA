VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmCadastroManual 
   Caption         =   "Cadastro Manual"
   ClientHeight    =   4335
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   6345
   OleObjectBlob   =   "FrmCadastroManual.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmCadastroManual"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Image1_Click()

Dim lSQL As String

If Me.txtData.Value = "" Then
MsgBox "Data Vazia!"
Else
On Error GoTo catch
 If Me.txtJust.Text = "" Or Len(Me.txtJust.Value) < 5 Or Len(Me.txtData.Value) < 10 Or Len(Me.txtHora.Value) < 8 Or CDate(Me.txtData.Value) >= vba.Date Then
          
         MsgBox "Campos Não Preenchidos ou Data inválida !", vbOKOnly, "Atenção"
         
         Else
                 IsInserir Me.txtData.Value, Environ$("username"), Environ$("computername"), Me.txtJust.Value, Me.txtHora.Value
                 
                 If Metodos.LimitInsert(CDate(Data)) = True Then
                     Call TEMP.InserirTotal(Login, Data)
                 End If
              
                     
    End If
  End If
catch:
    With frmCadastroManual
    .txtData.Value = ""
    .txtHora = ""
    .txtJust = ""
    End With
End Sub

Private Sub txtData_Change()

    If Len(txtData.Value) = 2 Or Len(txtData.Value) = 5 Then
    
        txtData.Value = txtData.Value & "/"
    
    End If
End Sub

Private Sub txtHora_Change()
       If Len(txtHora.Value) = 2 Or Len(txtHora.Value) = 5 Then
    
        txtHora.Value = txtHora.Value & ":"
    
       End If

End Sub



