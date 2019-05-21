VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmTipo 
   ClientHeight    =   7245
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10260
   OleObjectBlob   =   "FrmTipo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmTipo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub btnAbono_Click()

Dim lSQL
Dim DiaFalta As Boolean
Dim Listbox
Listbox = LBConsulta.List(Me.LBConsulta.ListIndex)
If Me.LBConsulta.Selected(Me.LBConsulta.ListIndex) = True Then

            SQL.lsConectar
            SQL.rs.Open "SELECT Data FROM HorasTotais WHERE Data =" & "#" & Day(Listbox) & "/" & Month(Listbox) & "/" & Year(Listbox) & "#", aux
            
            If SQL.rs.BOF Then
                DiaFalta = True
            End If
            
            SQL.rs.Close
            SQL.lsDesconectar


            If DiaFalta = False Then
                SQL.lsConectar
                lSQL = "UPDATE HorasTotais "
                lSQL = lSQL & " SET Situacao ='"
                lSQL = lSQL & Me.cbAbono.Value & "'"
                
                If Me.cbAbono.Value = "Férias" Then
                        lSQL = lSQL & ",Abono=" & "#" & LBConsulta.List(Me.LBConsulta.ListIndex, 4) & "# " ' Total Horas Abonadas
                ElseIf Me.cbAbono.Value = "Atestado" Then
                        lSQL = lSQL & ",Abono =" & "'" & CDate(Me.txtHorasAbono.Value) & "'" ' Total Horas Abonadas
                End If
                
                    lSQL = lSQL & " WHERE LoginServer ='" & txtLogin.Value & "'" & " AND Data =" & "#" & Day(Listbox) & "/" & Month(Listbox) & "/" & Year(Listbox) & "#"
                    SQL.gConexao.Execute lSQL
                    SQL.lsDesconectar
                    
            Else
            
                'Inserir Abono na Tabela "HorasTotais"
                SQL.lsConectar
                ISQL = ISQL & "INSERT INTO HorasTotais(LoginServer,Data,Hora,Situacao,Abono) "
                ISQL = ISQL & "VALUES('" & Me.txtLogin.Value & "'," 'LoginServer
                ISQL = ISQL & "#" & Month(Listbox) & "/" & Day(Listbox) & "/" & Year(Listbox) & "#" & ","    'Data
                ISQL = ISQL & "'" & CDate("00:00:00") & "','"       ' HoraTotal
                ISQL = ISQL & Me.cbAbono & "'," 'Situação
                If Me.cbAbono.Value = "Férias" Then
                    ISQL = ISQL & "'" & Right(LBConsulta.List(Me.LBConsulta.ListIndex, 2), 8) & "')" ' Abono
                ElseIf Me.cbAbono.Value = "Atestado" Then
                    ISQL = ISQL & "'" & CDate(Me.txtHorasAbono.Value) & "')" ' Abono
                End If
                SQL.gConexao.Execute ISQL
                SQL.lsDesconectar
                
                'ATT(Ajuste Técnico Temporário)
                Call SQL.IsInserir(Day(Listbox) & "/" & Month(Listbox) & "/" & Year(Listbox), Me.txtLogin.Value, "", Me.cbAbono.Value, "")
                
            End If
            
       'Query para puxar o nome do Funcionário
       SQL.lsConectar
       SQL.rs.Open "SELECT INFO.Nome  " _
       & " FROM INFO " _
       & " INNER JOIN HorasTotais ON INFO.LoginServer = HorasTotais.LoginServer " _
       & " WHERE   INFO.LoginServer= " & "'" & txtLogin.Value & "'", aux
                                      
       Do Until SQL.rs.EOF
            Matriz = SQL.rs.GetRows
       Loop
       
            SQL.lsDesconectar
            SQL.rs.Close
        Visualizar Me.cbMes.ListIndex + 1, Matriz(0, 0)
        
            
End If

End Sub

Private Sub btnExtra_Click()
Dim lSQL

If Me.LBConsulta.Selected(Me.LBConsulta.ListIndex) = True Then
    
    If Left(LBConsulta.List(Me.LBConsulta.ListIndex, 2), 1) <> "-" Then
        
         SQL.lsConectar
                lSQL = "UPDATE HorasTotais "
                lSQL = lSQL & " SET Extra ="
                lSQL = lSQL & "'" & Me.cbTipo.Value & "' "
                lSQL = lSQL & " WHERE LoginServer ='" & txtLogin.Value & "'" & " AND Data =" & "#" & vba.Format(LBConsulta.List(Me.LBConsulta.ListIndex), "dd/mm/yyyy") & "#"
                SQL.gConexao.Execute lSQL
                SQL.lsDesconectar
    End If
    
    'Query para puxar o nome do Funcionário
       SQL.lsConectar
       SQL.rs.Open "SELECT INFO.Nome  " _
       & " FROM INFO " _
       & " INNER JOIN HorasTotais ON INFO.LoginServer = HorasTotais.LoginServer " _
       & " WHERE   INFO.LoginServer= " & "'" & txtLogin.Value & "'", aux
                                      
       Do Until SQL.rs.EOF
            Matriz = SQL.rs.GetRows
       Loop
            SQL.lsDesconectar
            SQL.rs.Close
        Me.Visualizar Me.cbMes.ListIndex + 1, Matriz(0, 0)
       
        
End If

End Sub

Private Sub cbAbono_Change()
    
    Select Case Me.cbAbono
        
        Case "Atestado"
            Me.lblAbono.Visible = True
            Me.txtHorasAbono.Visible = True
        Case ""
        Me.lblAbono.Visible = False
        Me.txtHorasAbono.Visible = False
        Case "Férias"
        Me.lblAbono.Visible = False
        Me.txtHorasAbono.Visible = False
    End Select

End Sub

Private Sub cbFunc_Change()
       'Call Visualizar(Me.cbMes.Value, Me.cbFunc.Value)
       Select Case Me.cbMes
        
        Case "janeiro"
           Call Visualizar(1, Me.cbFunc.Value)
            
            Case "fevereiro"
           Call Visualizar(2, Me.cbFunc.Value)
            
            Case "março"
            Call Visualizar(3, Me.cbFunc.Value)
            
            Case "abril"
           Call Visualizar(4, Me.cbFunc.Value)
            
            Case "maio"
            Call Visualizar(5, Me.cbFunc.Value)
            
            Case "junho"
            Call Visualizar(6, Me.cbFunc.Value)
            
            Case "julho"
            Call Visualizar(7, Me.cbFunc.Value)
            
            Case "agosto"
             Call Visualizar(8, Me.cbFunc.Value)
            
            Case "setembro"
           Call Visualizar(9, Me.cbFunc.Value)
            
            Case "outubro"
            Call Visualizar(10, Me.cbFunc.Value)
            
            Case "novembro"
            Call Visualizar(11, Me.cbFunc.Value)
            
            Case "dezembro"
            Call Visualizar(12, Me.cbFunc.Value)
       
       End Select
End Sub



Private Sub cbMes_Change()
Select Case Me.cbMes
        
        Case "janeiro"
           Call Visualizar(1, Me.cbFunc.Value)
            
            Case "fevereiro"
           Call Visualizar(2, Me.cbFunc.Value)
            
            Case "março"
            Call Visualizar(3, Me.cbFunc.Value)
            
            Case "abril"
           Call Visualizar(4, Me.cbFunc.Value)
            
            Case "maio"
            Call Visualizar(5, Me.cbFunc.Value)
            
            Case "junho"
            Call Visualizar(6, Me.cbFunc.Value)
            
            Case "julho"
            Call Visualizar(7, Me.cbFunc.Value)
            
            Case "agosto"
             Call Visualizar(8, Me.cbFunc.Value)
            
            Case "setembro"
           Call Visualizar(9, Me.cbFunc.Value)
            
            Case "outubro"
            Call Visualizar(10, Me.cbFunc.Value)
            
            Case "novembro"
            Call Visualizar(11, Me.cbFunc.Value)
            
            Case "dezembro"
            Call Visualizar(12, Me.cbFunc.Value)
       
       End Select
End Sub


Private Sub rbAbono_Click()
    Me.cbTipo.Visible = False
    Me.lbl19.Visible = False
    Me.cbAbono.Visible = True
    Me.lbl15.Visible = True
    Me.btnAbono.Visible = True
    Me.btnExtra.Visible = False
End Sub



Private Sub rbExtra_Click()
    Me.cbTipo.Visible = True
    Me.lbl19.Visible = True
    
    Me.cbAbono.Visible = False
    Me.lbl15.Visible = False
    Me.btnAbono.Visible = False
    btnExtra.Visible = False
    Me.txtHorasAbono.Visible = False
    lblAbono.Visible = False
    Me.btnExtra.Visible = True
    
End Sub

Public Sub UserForm_Initialize()
                  
    Dim MatrizNome As Variant
    Me.txtLogin = Null
        
    With Me.cbTipo
        .AddItem "Hora-Extra", 0
        .AddItem "Banco de Horas", 1
    End With
    
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
    
    With Me.cbAbono
        .AddItem "Férias", 0
        .AddItem "Atestado", 1
    End With
      
   SQL.lsConectar
   SQL.rs.Open "SELECT  Nome,Periodo,LoginServer " _
   & " From INFO ", SQL.aux
   MatrizNome = SQL.rs.GetRows
   SQL.rs.Close
   SQL.lsDesconectar
      
     With Me
        For Contador = 0 To UBound(MatrizNome, 2)
            cbFunc.AddItem MatrizNome(0, Contador), 0
        Next Contador
         
    End With
    
     cbFunc.Value = vba.CStr(MatrizNome(0, 0))
     txtLogin.Value = MatrizNome(2, 0)
     cbMes.Text = vba.Format(vba.DateTime.Now, "mmmm")
     
End Sub

  Public Function Visualizar(texto, nome)
  
                Me.txtNegt = Null
                Dim Total As Variant
                Dim LinhaListbox  As Integer
                Dim LinhaHora As Integer
                Dim ColunaHora As Integer
                Dim Contador As Integer
                Dim Matriz As Variant
                Dim MatrizDias As Variant
                Dim strSQL As String
                Dim NomeFucionario As String
                
                Matriz = Null
                LinhaListbox = 0
                Coluna = 0
                
                With Me.LBConsulta
                    .Clear
                    .ColumnCount = 6 'Número de colunas no Listbox
                    .ColumnWidths = "60;50;70;60;80;60"
                End With
                
                SQL.lsConectar
                strSQL = strSQL & "SELECT LoginServer FROM INFO WHERE Nome = '" & Me.cbFunc.Value & "'"
                SQL.rs.Open strSQL, aux
                txtLogin.Value = SQL.rs!LoginServer
                SQL.lsDesconectar
                SQL.rs.Close
                
                strSQL = ""
                         
                'Query para obter dados da tabela de horários já salvos  com a tabela que contêm as demais informações
               SQL.lsConectar
                  strSQL = strSQL & "SELECT  "
                  strSQL = strSQL & "D.Data, "
                  strSQL = strSQL & "B.LoginServer, "
                  strSQL = strSQL & "B.Hora, "
                  strSQL = strSQL & "I.Horas_seg_qui, "
                  strSQL = strSQL & "I.Horas_sex, "
                  strSQL = strSQL & "HorasTotais.Situacao, "
                  strSQL = strSQL & "HorasTotais.Abono, "
                  strSQL = strSQL & "HorasTotais.Extra "
                  strSQL = strSQL & "FROM (((Dias D "
                  strSQL = strSQL & "LEFT JOIN (select LoginServer, Hora, Data from Base where  LoginServer = '" & txtLogin.Value & "') B ON D.Data = B.Data) "
                  strSQL = strSQL & "LEFT JOIN INFO I ON IIF(B.LoginServer IS NULL," & "'" & txtLogin.Value & "'" & ",B.LoginServer) = I.LoginServer) "
                  strSQL = strSQL & "LEFT JOIN HorasTotais ON B.LoginServer = HorasTotais.LoginServer AND B.Data = HorasTotais.Data) "
                  strSQL = strSQL & "WHERE "
                  
                  If CInt(Me.cbMes.ListIndex + 1) = CInt(Month(Now)) Then
                    strSQL = strSQL & "MONTH(D.Data) = '" & texto & "' AND D.Dia = 'DiaUtil' AND D.Data <= #" & vba.Date & "#" & " AND "
                  ElseIf CInt(Me.cbMes.ListIndex + 1) < CInt(Month(Now)) Then
                    strSQL = strSQL & "MONTH(D.Data) = '" & texto & "' AND D.Dia = 'DiaUtil'  AND "
                  ElseIf CInt(Me.cbMes.ListIndex + 1) > CInt(Month(Now)) Then
                    Exit Function
                    Me.txtNegt = Null
                  End If
                  
                  strSQL = strSQL & "(B.LoginServer =" & "'" & txtLogin.Value & "' OR "
                  strSQL = strSQL & " B.LoginServer Is Null) "
                  strSQL = strSQL & "ORDER BY D.Data,B.Hora "
                  SQL.rs.Open strSQL, aux
               
                    Do Until SQL.rs.EOF
                            'MATRIZ -> Matriz que recebe dados da Query acima
                            Matriz = SQL.rs.GetRows
                            numero_de_Registros = UBound(Matriz, 2)
                            Matr = Matriz
                    Loop
                            SQL.lsDesconectar
                            SQL.rs.Close 'Fecha conexão recordset
                                            
                            
                            For Contador = 0 To numero_de_Registros
                            
                               On Error GoTo Handler 'Tratamento de erro caso o mês selecionado não exista
                               
                                            With Me.LBConsulta
                                                If Contador = 0 Then
                                                    .AddItem
                                                ElseIf Matriz(0, Contador) <> Matriz(0, Contador - 1) Then
                                                    If Matriz(0, Contador) > vba.Date Then
                                                        Me.txtNegt.Value = Me.Somar(Me.LBConsulta, Matr)
                                                        Exit Function
                                                    Else
                                                        .AddItem
                                                    End If
                                                End If
                                                                                                
                                                .Column(0, LinhaListbox) = Matriz(0, Contador) 'Data
                                                 
                                                 Total = TEMP.ContarHoras(CDate(Matriz(0, Contador)), txtLogin.Value)
                                                .Column(1, LinhaListbox) = Total 'Total de Horas no dia
                                                
                                                'Verifica se o dia da semana está entre segunda a quinta-feira
                                                If vba.Format(CDate(Matriz(0, Contador)), "dddd") <> "sexta-feira" Then
                                                
                                                    If CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(3, Contador))) < 0 Then 'Veifica horas negativas
                                                        .Column(2, LinhaListbox) = "-" & CDate(CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(3, Contador)))) 'Concatena "-" caso seja horas negativas
                                                    Else
                                                        .Column(2, LinhaListbox) = " " & CDate(CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(3, Contador))))
                                                    End If
                                                    
                                                Else 'verifica se o dia da semana é sexta-feira
                                                
                                                   If CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(4, Contador))) < 0 Then 'Verifica horas negativas
                                                        .Column(2, LinhaListbox) = "-" & CDate(CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(4, Contador)))) 'Concatena "-" caso seja horas negativas
                                                    Else
                                                        .Column(2, LinhaListbox) = " " & CDate(CDbl(vba.TimeValue(Total) - vba.TimeValue(Matriz(4, Contador))))
                                                    End If
                                                    
                                                End If
                                                
                                                 If Matriz(5, Contador) <> "" Then
                                                    .Column(3, LinhaListbox) = Matriz(5, Contador)
                                                End If
                                                
                                                   ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''Férias
                                                If Matriz(5, Contador) = "Férias" Then
                                                        If vba.Format(CDate(Matriz(0, Contador)), "dddd") <> "sexta-feira" Then
                                                            .Column(4, LinhaListbox) = Matriz(3, Contador)
                                                        Else
                                                            .Column(4, LinhaListbox) = Matriz(4, Contador)
                                                        End If
                                                End If
                                                ''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
                                                
                                               '''''''''''''''''''''''''''''''''''''''''''''''Atestado
                                                 If Matriz(5, Contador) = "Atestado" Then
                                                        If Me.txtHorasAbono.Value = "" And Matriz(6, Contador) <> "" Then
                                                                     .Column(4, LinhaListbox) = Matriz(6, Contador)
                                                        Else
                                                                     .Column(4, LinhaListbox) = TimeValue(Me.txtHorasAbono)
                                                        End If
                                                End If
                                                '''''''''''''''''''''''
                                                  '''''''''''''''''''''''''''''''''''''''''''''''Normal
                                                 If IsNull(Matriz(5, Contador)) Or Matriz(5, Contador) = "" Then
                                                                     .Column(4, LinhaListbox) = CDate("00:00:00")
                                                End If
                                                '''''''''''''''''''''''
                                                
                                                .Column(5, LinhaListbox) = IIf(IsNull(Matriz(7, Contador)), "", Matriz(7, Contador))
                                                
                                            End With
                                       
                                      If Contador <> numero_de_Registros Then
                                             'Verifica se a Data atual é difente da próxima linha da matriz
                                             If Matriz(0, Contador) <> Matriz(0, Contador + 1) Then
                                                    LinhaListbox = LinhaListbox + 1
                                            End If
                                    End If
                                Next
                                Me.txtNegt.Value = Me.Somar(Me.LBConsulta, Matr)
Handler:
    End Function

Public Function Somar(LB As MSForms.Listbox, Matriz As Variant) As String

        Dim lItem As Double
        Dim Total As String
        Dim HoraPositiva As Double
        Dim HoraNegativa As Double
        
        If LB.ListCount = 0 Then
            Exit Function
        End If
        
        For lItem = 0 To LB.ListCount - 1
            
            If Left(LB.List(lItem, 2), 1) = "-" Then
            
                 If LB.List(lItem, 3) = "Férias" Then
                 
                    HoraNegativa = 0
                    
                 ElseIf LB.List(lItem, 3) = "Atestado" Or IsNull(LB.List(lItem, 3)) Then
                    HoraNegativa = HoraNegativa + CDbl(CDate(Right(LB.List(lItem, 2), 8)))
                    HoraNegativa = HoraNegativa - CDbl(CDate(Right(LB.List(lItem, 4), 8)))
                 End If
            
                
            ElseIf Left(LB.List(lItem, 2), 1) <> "-" Then
                        HoraPositiva = HoraPositiva + CDbl(CDate(Right(LB.List(lItem, 2), 8)))
                        
            End If
        
        Next
        
            If HoraPositiva > HoraNegativa Then
                Total = Metodos.FormatHM(HoraPositiva - HoraNegativa)
            ElseIf HoraPositiva < HoraNegativa Then
                Total = "-" & Metodos.FormatHM(HoraNegativa - HoraPositiva)
                
    End If


Somar = Total
End Function
