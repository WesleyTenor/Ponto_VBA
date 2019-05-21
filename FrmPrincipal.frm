VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmPrincipal 
   Caption         =   "Cadastro"
   ClientHeight    =   7035
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8295.001
   OleObjectBlob   =   "FrmPrincipal.frx":0000
   StartUpPosition =   1  'CenterOwner
   WhatsThisHelp   =   -1  'True
End
Attribute VB_Name = "FrmPrincipal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnMostrar_Click()
    Application.Visible = True
    Unload Me
End Sub

Private Sub cbMes_Change()

 
    Select Case Me.cbMes
        
            Case "janeiro"
            SQL.Flitrar 1
            
            Case "fevereiro"
            SQL.Flitrar 2
            
            Case "março"
            SQL.Flitrar 3
            
            Case "abril"
            SQL.Flitrar 4
            
            Case "maio"
            SQL.Flitrar 5
            
            Case "junho"
            SQL.Flitrar 6
            
            Case "julho"
            SQL.Flitrar 7
            
            Case "agosto"
            SQL.Flitrar 8
            
            Case "setembro"
            SQL.Flitrar 9
            
            Case "outubro"
            SQL.Flitrar 10
            
            Case "novembro"
            SQL.Flitrar 11
            
            Case "dezembro"
            SQL.Flitrar 12
       
       End Select
       
    
End Sub


Private Sub Image1_Click()
FrmTipo.Show
End Sub


Private Sub ImgInsert_Click()
SQL.IsInserir vba.Date, Environ$("username"), Environ$("computername"), "", vba.Time
End Sub


Private Sub UserForm_Initialize()
Me.btnMostrar.BackColor = RGB(0, 216, 150)
    
    Dim CorLabel
    CorLabel = vba.RGB(0, 216, 150)
    
    Me.Label3.ForeColor = CorLabel
    Me.Label4.ForeColor = CorLabel
    Me.Label5.ForeColor = CorLabel
    Me.Label6.ForeColor = CorLabel
    
    frmCadastroManual.Label4.ForeColor = CorLabel
    frmCadastroManual.Label5.ForeColor = CorLabel
    frmCadastroManual.Label6.ForeColor = CorLabel
    FrmPrincipal.BackColor = vba.RGB(15, 13, 89)
    
    
    Me.cbMes.Value = vba.Format(vba.DateTime.Now, "mmmm")
    SQL.Flitrar (vba.Format(vba.DateTime.Now, "m"))

    'Preencher Combo Box
    With Me.cbMes
        .AddItem "janeiro", 0
        .AddItem "fevereiro", 1
        .AddItem "março", 2
        .AddItem "abril", 3
        .AddItem "maio", 4
        .AddItem "junho", 5
        .AddItem "julho", 6
        .AddItem "agosto", 7
        .AddItem "setembro", 8
        .AddItem "outubro", 9
        .AddItem "novembro", 10
        .AddItem "dezembro", 11
    End With
 
End Sub


Private Sub UserForm_Terminate()
    If Application.Visible = False Then
        Application.Quit
    End If
End Sub
