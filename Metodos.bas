Attribute VB_Name = "Metodos"
'Método que verifica se a data existe no banco
Public Function Valid_Insert() As Boolean
        On Error GoTo Handler
        Dim Matriz As Variant
                      
       SQL.lsConectar
       SQL.rs.Open "SELECT Data,COUNT(Data) as Teste " _
       & "FROM Base " _
       & "WHERE Data = " & CDbl(CDate(Date)) & "AND LoginServer = " & "'" & Environ$("username") & "'" _
       & " GROUP BY Data", aux
       Do Until SQL.rs.EOF
       Matriz = SQL.rs.GetRows
        Loop
        
        SQL.lsDesconectar
        SQL.rs.Close
        
        If Matriz = Empty Then
             Valid_Insert = True
        End If
        Exit Function
Handler:
Valid_Insert = False
End Function

'Método que impede de inserir mais de 4 horários em uma única data
Public Function LimitInsert(Data As String) As Boolean
On Error GoTo Handler
 Dim Matriz As Variant
                      
       SQL.lsConectar
       SQL.rs.Open "SELECT Data,COUNT(Data) as Teste FROM Base   WHERE Data = " & CDbl(CDate(Data)) & " GROUP BY Data", aux
       Do Until SQL.rs.EOF
       Matriz = SQL.rs.GetRows
       Loop
        SQL.lsDesconectar
        SQL.rs.Close
        
         If Matriz(1, 0) = 4 Then
               LimitInsert = True
        Else
                LimitInsert = False
         End If
         Exit Function
Handler:
  LimitInsert = False
End Function

Public Function FormatHM(v As Double) As String
    FormatHM = Format(Application.Floor(v * 24, 1), "00") & _
                 ":" & Format((v * 1440) Mod 60, "00")
End Function

Public Function HrStr(dblHora As Double) As String

Dim strHoras As String
Dim strMinutos As String

'Pega as horas (parte inteira)
strHoras = CStr(Fix(dblHora))

'Pega os minutos
strMinutos = vba.Format$(Abs((dblHora - vba.Fix(dblHora)) * 60), "00")

'Verifica se o total de minutos é 60
If strMinutos = "60" Then
strMinutos = "00"
strHoras = CStr(CDbl(strHoras) + 1)
End If

'Concatena os dois
HrStr = strHoras & ":" & strMinutos


End Function

