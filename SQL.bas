Attribute VB_Name = "SQL"
Public gConexao As New ADODB.Connection
Public rs As New ADODB.Recordset

Global aux As String 'Global contendo a string de conexão
Global Matr As Variant

Public Function lsConectar()
    Dim strConexao As String
    Dim lCaminho   As String
    Set gConexao = New ADODB.Connection
    'String do caminho do banco de dados
    lCaminho = Replace(ThisWorkbook.FullName, ThisWorkbook.Name, "")
    lCaminho = lCaminho & "Base.accdb"
    'String de conexão OLEDB
    strConexao = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" & lCaminho & ";Persist Security Info=False"
    aux = strConexao
    gConexao.Open strConexao
End Function

'Desconectar do acess
Public Function lsDesconectar()
    If Not gConexao Is Nothing Then
        gConexao.Close
        Set gConexao = Nothing
    End If
End Function

'Método que insere os dados no acess
Public Function IsInserir(Data, Login, hostname, justf, hora)
        Dim lSQL As String
        'Verificações antes de adicionar dados no banco de dados "Acess"
        If Metodos.Valid_Insert = True Then
            SQL.lsConectar
            lSQL = "INSERT INTO Base (Data, LoginServer,HostName,Justificativa,Hora )" & _
            "VALUES ( """ & Data & """,""" & Login & """,""" & hostname & """,""" & justf & """ ,""" & hora & """ )"
            SQL.gConexao.Execute lSQL
            lsDesconectar
            SQL.Flitrar (vba.Format(vba.DateTime.Now, "m"))
            
            If Metodos.LimitInsert(CDate(Data)) = True Then
                     Call TEMP.InserirTotal(Login, Data)
            End If
            
        ElseIf Metodos.Valid_Insert = False Then
            If Metodos.LimitInsert(CDbl(CDate(Data))) = True Then
            MsgBox "4 Horários para essa data já preenchidos!", vbOKOnly, "Atenção"
            SQL.Flitrar (vba.Format(vba.DateTime.Now, "m"))
            Else
            SQL.lsConectar
            lSQL = "INSERT INTO Base (Data, LoginServer,HostName,Justificativa,Hora )" & _
            "VALUES ( """ & Data & """,""" & Login & """,""" & hostname & """,""" & justf & """ ,""" & hora & """ )"
            SQL.gConexao.Execute lSQL
            SQL.lsDesconectar
            SQL.Flitrar (vba.Format(vba.DateTime.Now, "m"))
            
            If Metodos.LimitInsert(CDate(Data)) = True Then
                 Call TEMP.InserirTotal(Login, Data)
            End If
            
            End If
        End If
        
End Function


    'Método para poupular Listbox na inicialização
    Public Function Flitrar(texto)
                Dim Total As Variant
                Dim LinhaListbox  As Integer
                Dim LinhaHora As Integer
                Dim ColunaHora As Integer
                Dim Contador As Integer
                Dim Matriz As Variant
                Dim MatrizDias As Variant
                Dim numero_de_Registros
                Dim strSQL As String
                
                Matriz = Null
                LinhaListbox = 0
                ColunaHora = 1
                LinhaHora = 0
                Coluna = 0
                FrmPrincipal.lblhorasT = ""
                strSQL = ""
                
                With FrmPrincipal.ListBox1
                    .Clear
                    .ColumnCount = 7 'Número de colunas no Listbox
                    .ColumnWidths = "60;50;50;50;50;50;50"
                End With
          
                  'Query para obter dados da tabela de horários já salvos  com a tabela que contêm as demais informações
                  SQL.lsConectar
                  strSQL = strSQL & "SELECT  "
                  strSQL = strSQL & "D.Data, "
                  strSQL = strSQL & "B.Hora, "
                  strSQL = strSQL & "B.LoginServer, "
                  strSQL = strSQL & "I.Horas_seg_qui, "
                  strSQL = strSQL & "I.Horas_sex, "
                  strSQL = strSQL & "HorasTotais.Situacao, "
                  strSQL = strSQL & "HorasTotais.Abono "
                  strSQL = strSQL & "FROM (((Dias D "
                  strSQL = strSQL & "LEFT JOIN (select LoginServer, Hora, Data from Base where  LoginServer = '" & Environ$("username") & "') B ON D.Data = B.Data) "
                  strSQL = strSQL & "LEFT JOIN INFO I ON IIF(B.LoginServer IS NULL," & "'" & Environ$("username") & "'" & ",B.LoginServer) = I.LoginServer) "
                  strSQL = strSQL & "LEFT JOIN HorasTotais ON B.LoginServer = HorasTotais.LoginServer AND B.Data = HorasTotais.Data) "
                  strSQL = strSQL & "WHERE "
                  
                  If CInt(FrmPrincipal.cbMes.ListIndex + 1) = CInt(Month(Now)) Then
                    strSQL = strSQL & "MONTH(D.Data) = '" & texto & "' AND D.Dia = 'DiaUtil'  AND D.Data <= #" & vba.Date & "# " & "AND "
                  ElseIf CInt(FrmPrincipal.cbMes.ListIndex + 1) < CInt(Month(Now)) Then
                    strSQL = strSQL & "MONTH(D.Data) = '" & texto & "' AND D.Dia = 'DiaUtil'  AND "
                  ElseIf CInt(FrmPrincipal.cbMes.ListIndex + 1) > CInt(Month(Now)) Then
                    Exit Function
                    FrmPrincipal.lblhorasT = Null
                  End If
                  
                  strSQL = strSQL & "(B.LoginServer =" & "'" & Environ$("username") & "' OR "
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
                               
                                            With FrmPrincipal.ListBox1
                                                If Contador = 0 Then
                                                    .AddItem
                                                ElseIf Matriz(0, Contador) <> Matriz(0, Contador - 1) Then
                                                    If Matriz(0, Contador) > vba.Date Then
                                                        IIf Contador - 1 = numero_de_Registros, TEMP.Somar(Matr), False
                                                    Else
                                                        .AddItem
                                                    End If
                                                End If
                                                .Column(0, LinhaListbox) = Matriz(0, Contador) 'Data
                                                .Column(ColunaHora, LinhaHora) = IIf(IsNull(Matriz(1, Contador)), "", Matriz(1, Contador))    'Hora
                                                Total = TEMP.ContarHoras(CDate(Matriz(0, Contador)), Environ$("username"))
                                                .Column(5, LinhaListbox) = Total 'Total de Horas no dia
                                                
                                                If IsNull(Matriz(6, Contador)) Then
                                                                                                
                                                  'Verifica se o dia da semana está entre segunda a quinta-feira
                                                  If vba.Format(CDate(Matriz(0, Contador)), "dddd") <> "sexta-feira" Then
                                                  
                                                        If CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(3, Contador)), TimeValue("00:00:00"), vba.TimeValue(Matriz(3, Contador)))) < 0 Then    'Verifica horas negativas
                                                            .Column(6, LinhaListbox) = "-" & CDate(CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(3, Contador)), "00:00:00", vba.TimeValue(Matriz(3, Contador))))) 'Concatena "-" caso seja horas negativas
                                                        Else
                                                            .Column(6, LinhaListbox) = " " & CDate(CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(3, Contador)), "00:00:00", vba.TimeValue(Matriz(3, Contador)))))
                                                        End If
                                                  
                                                  Else 'verifica se o dia da semana é sexta-feira
                                                  
                                                        If CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(4, Contador)), TimeValue("00:00:00"), vba.TimeValue(Matriz(4, Contador)))) < 0 Then 'Verifica horas negativas
                                                         .Column(6, LinhaListbox) = "-" & CDate(CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(4, Contador)), TimeValue("00:00:00"), vba.TimeValue(Matriz(4, Contador))))) 'Concatena "-" caso seja horas negativas
                                                         Else
                                                         .Column(6, LinhaListbox) = " " & CDate(CDbl(vba.TimeValue(Total) - IIf(IsNull(Matriz(4, Contador)), TimeValue("00:00:00"), vba.TimeValue(Matriz(4, Contador)))))
                                                        End If
                                                   
                                                   End If
                                                  Else
                                                  
                                                        If Matriz(5, Contador) = "Atestado" Then
                                                         Resto = CDbl(CDate(Matriz(6, Contador))) + CDbl(CDate(FrmPrincipal.ListBox1.List(LinhaListbox, 5)))
                                                             .Column(6, LinhaListbox) = CDate(Resto)
                                                        Else
                                                            .Column(6, LinhaListbox) = Matriz(6, Contador)
                                                        End If
                                                         
                                                End If
                                                
                                            End With
                                      
                                      If Contador <> numero_de_Registros Then
                                             'Verifica se a Data atual é difente da próxima linha da matriz
                                             If Matriz(0, Contador) <> Matriz(0, Contador + 1) Then
                                                    LinhaListbox = LinhaListbox + 1
                                                    LinhaHora = LinhaHora + 1
                                                    ColunaHora = 1
                                                    'Se a data for a mesma permanece na mesma linha (Listbox)
                                                    ElseIf Matriz(0, Contador) = Matriz(0, Contador + 1) Then
                                                    ColunaHora = ColunaHora + 1
                                            End If
                                    End If
                                      
                                Next
                                
                                    
Handler:
      IIf Contador - 1 = numero_de_Registros, TEMP.Somar(Matr), False
    End Function
     

