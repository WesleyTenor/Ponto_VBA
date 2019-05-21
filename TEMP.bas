Attribute VB_Name = "TEMP"
Public rs As New ADODB.Recordset
Public gConexao As New ADODB.Connection

Public Function Somar(var As Variant) As String  '24/01/2019
Dim lItem As Double
Dim Total As String
Dim HoraPositiva As Double
Dim HoraNegativa As Double

For lItem = 0 To FrmPrincipal.ListBox1.ListCount - 1

    If IsEmpty(var) Then
        Exit Function
    End If

    If CDate(FrmPrincipal.ListBox1.List(lItem, 0)) <= vba.Date Then 'Data (Matriz)
        
                If FrmPrincipal.cbMes.Value = Format(CDate(var(0, lItem)), "mmmm") Then
                
                    If Left(FrmPrincipal.ListBox1.List(lItem, 6), 1) = "-" Then
                        HoraNegativa = HoraNegativa + CDbl(CDate(Right(FrmPrincipal.ListBox1.List(lItem, 6), 8)))
                      
                    ElseIf Left(FrmPrincipal.ListBox1.List(lItem, 6), 1) <> "-" Then
                                HoraPositiva = HoraPositiva + CDbl(CDate(Right(FrmPrincipal.ListBox1.List(lItem, 6), 8)))
                    End If
                End If
    End If
            Next
            
                If HoraPositiva > HoraNegativa Then
                    Total = Metodos.FormatHM(HoraNegativa)
                ElseIf HoraPositiva < HoraNegativa Then
                    Total = "-" & Metodos.FormatHM(HoraNegativa)
                End If

    
FrmPrincipal.lblhorasT = Total
End Function





Public Function ContarHoras(Data As String, Login) As String
            
                Dim Primeira As Variant
                Dim Segunda As Variant
                Dim Total As Variant
                Dim ISQL
                
                    
                    SQL.lsConectar
                     SQL.rs.Open "SELECT Hora,Data " _
                    & "FROM Base " _
                    & "WHERE Data = #" & CDate(Data) & "# AND LoginServer = " & "'" & Login & "'" _
                    & "GROUP BY Data,Hora", aux
                    
                    On Error GoTo jump
                    SQL.rs.Move 1 'Saída Almoço
                    Primeira = Primeira + CDbl(CDate(SQL.rs![hora]))
                  
                    SQL.rs.MoveFirst 'Entrada
                    Primeira = Primeira - CDbl(CDate(SQL.rs![hora]))
                  
                    SQL.rs.Move 2 'Volta Almoço
                    Segunda = Segunda + CDbl(CDate(SQL.rs![hora]))
                    
                    SQL.rs.MoveNext 'Saída
                     Segunda = Segunda - CDbl(CDate(SQL.rs![hora]))
                    
jump:
                    Total = Total + (CDate(Primeira) - CDate(Segunda))
             
           ' ContarHoras = IIf(Total = 0, "", CDate(Total))
                ContarHoras = CDate(Total)
                SQL.rs.Close
                SQL.lsDesconectar
                
End Function



Private Function Alter()
    Dim ISQL
    Dim var
    SQL.lsConectar
    ISQL = "ALTER TABLE HorasTotais " _
    & "ADD COLUMN Extra Varchar(255) "
    SQL.gConexao.Execute ISQL
    SQL.lsDesconectar
End Function



Public Sub lsCreate()
            Dim lSQL
           SQL.lsConectar
            lSQL = "CREATE TABLE HorasTotais(LoginServer varchar(255),Data datetime, Hora datetime, Situacao VARCHAR(255))"
          '  lSQL = "DROP TABLE HorasTotais"
            SQL.gConexao.Execute lSQL
End Sub

Public Sub lsExcluir()
            Dim lSQL
           SQL.lsConectar
            lSQL = "DELETE FROM HorasTotais"
            SQL.gConexao.Execute lSQL
End Sub


Function InserirTotal(Login, Data)
           Dim lSQL As String
           Dim TotalHoras
           TotalHoras = TEMP.ContarHoras(CStr(Data), Login)
           SQL.lsConectar
           lSQL = "INSERT INTO HorasTotais(LoginServer,Data,Hora )" & _
          "VALUES ( """ & Login & """, #" & CDate(Data) & "# ,""" & TotalHoras & """ )"
           SQL.gConexao.Execute lSQL
           SQL.lsDesconectar
End Function



