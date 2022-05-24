Option Compare Database

Public Type sctPatrimonio
    Data As Date
    Assessor As String
    PL As Double
    Captacao As Double
    Deposito As Double
    Retirada As Double
    Net As Double
    Transf_Entrada As Double
    Transf_Saida As Double
    Transf_Total As Double
End Type

Function Atualiza_Dados()

Dim dbDados As DAO.Database
Dim rsAssessor As DAO.Recordset
Dim strConsulta As String
Dim Patrimonio() As sctPatrimonio
Dim intAssessor As Integer
Dim dtDataInicio, dtDataFinal As Date
Dim strAssessor As String
Dim rsCaptacao, rsPatrimonio As DAO.Recordset

dtDataAtualiza = "17/05/2022"
'Defini data inicio do mes
dtDataInicio = CDate(Format("01/" & Month(dtDataAtualiza) & "/" & Year(dtDataAtualiza), "dd/mm/yyyy"))
dtDataFinal = "31/05/2022"



'Defini banco
Set dbDados = CurrentDb
'Efetua consulta dos assesssores
strConsulta = "Select * from Assessores where Ativo=True order by Nome asc"
Set rsAssessor = dbDados.OpenRecordset(strConsulta)
ReDim Patrimonio(rsAssessor.RecordCount)
intAssessor = 1
Patrimonio(0).Assessor = "11120"
Patrimonio(0).Data = dtDataInicio

'Verifica se há dados no banco de diversificação
strConsulta = "Select MAX(Data) as MaxData from Diversificacao where Data>=#" & Format(dtDataInicio, "mm/dd/yyyy") & "# and Data<=#" & Format(dtDataFinal, "mm/dd/yyyy") & "#"
Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
If IsNull(rsCaptacao("Maxdata")) Then
    MsgBox "Não é possível fazer atualização dos dados. Não existe dados da diversificação para o calculo do PL dos assessores", vbCritical
    Exit Function
Else
    dtDataPL = rsCaptacao("Maxdata")
End If
    


While rsAssessor.EOF = False

    strAssessor = rsAssessor("Key_Codigo")
    Patrimonio(intAssessor).Assessor = strAssessor
    Patrimonio(intAssessor).Data = dtDataInicio

    '###################################################################
    'Verifica a quantidade de aporte financeiro do mes
    strConsulta = "Select Sum(Captacao) as ValorTotal from Captacao INNER JOIN Clientes On Captacao.codCliente = Clientes.codCliente Where Clientes.AssessorAtual='" & strAssessor & "' and Aux='C' and (Data>=#" & Format(dtDataInicio, "mm/dd/yyyy") & "# and Data<=#" & Format(dtDataFinal, "mm/dd/yyyy") & "#)"
    Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
    If IsNull(rsCaptacao("ValorTotal")) Then
        Patrimonio(intAssessor).Deposito = 0
    Else
        Patrimonio(intAssessor).Deposito = rsCaptacao("ValorTotal")
    End If

    'Verifica a quantidade de retirada financeiro do mes
    strConsulta = "Select Sum(Captacao) as ValorTotal from Captacao INNER JOIN Clientes On Captacao.codCliente = Clientes.codCliente Where Clientes.AssessorAtual='" & strAssessor & "' and Aux='D' and (Data>=#" & Format(dtDataInicio, "mm/dd/yyyy") & "# and Data<=#" & Format(dtDataFinal, "mm/dd/yyyy") & "#)"
    Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
    If IsNull(rsCaptacao("ValorTotal")) Then
        Patrimonio(intAssessor).Retirada = 0
    Else
        Patrimonio(intAssessor).Retirada = rsCaptacao("ValorTotal")
    End If
    
    Patrimonio(intAssessor).Net = Patrimonio(intAssessor).Deposito + Patrimonio(intAssessor).Retirada
   
    '###################################################################
    'Verifica a quantidade de TRANSFERENCIA DE CLIENTES (ENTRADA) do mes
    strConsulta = "Select Sum(PL) as ValorTotal from Entrada_Saida_Contas INNER JOIN Clientes On Entrada_Saida_Contas.codCliente = Clientes.codCliente Where Clientes.AssessorAtual='" & strAssessor & "' and Tipo='E' and (DataTransferencia>=#" & Format(dtDataInicio, "mm/dd/yyyy") & "# and DataTransferencia<=#" & Format(dtDataFinal, "mm/dd/yyyy") & "#)"
    Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
    If IsNull(rsCaptacao("ValorTotal")) Then
        Patrimonio(intAssessor).Transf_Entrada = 0
    Else
        Patrimonio(intAssessor).Transf_Entrada = rsCaptacao("ValorTotal")
    End If

    'Verifica a quantidade de TRANSFERENCIA DE CLIENTES (SAIDA)  do mes
    strConsulta = "Select Sum(PL) as ValorTotal from Entrada_Saida_Contas INNER JOIN Clientes On Entrada_Saida_Contas.codCliente = Clientes.codCliente Where Clientes.AssessorAtual='" & strAssessor & "' and Tipo='S' and (DataTransferencia>=#" & Format(dtDataInicio, "mm/dd/yyyy") & "# and DataTransferencia<=#" & Format(dtDataFinal, "mm/dd/yyyy") & "#)"
    Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
    If IsNull(rsCaptacao("ValorTotal")) Then
        Patrimonio(intAssessor).Transf_Saida = 0
    Else
        Patrimonio(intAssessor).Transf_Saida = -rsCaptacao("ValorTotal")
    End If
    Patrimonio(intAssessor).Transf_Total = Patrimonio(intAssessor).Transf_Entrada + Patrimonio(intAssessor).Transf_Saida
       
    Patrimonio(intAssessor).Captacao = Patrimonio(intAssessor).Net + Patrimonio(intAssessor).Transf_Total
    
    
    'Verifica o PL do Assessor
    strConsulta = "Select Sum(Net) as ValorTotal from Diversificacao INNER JOIN Clientes On Diversificacao.codCliente = Clientes.codCliente Where Clientes.AssessorAtual='" & strAssessor & "' and Data=#" & Format(dtDataPL, "mm/dd/yyyy") & "#"
    Set rsCaptacao = dbDados.OpenRecordset(strConsulta)
    If IsNull(rsCaptacao("ValorTotal")) Then
        Patrimonio(intAssessor).PL = 0
    Else
        Patrimonio(intAssessor).PL = rsCaptacao("ValorTotal")
    End If
    
    

    'Totaliza os dados da Invicta
    Patrimonio(0).Deposito = Patrimonio(0).Deposito + Patrimonio(intAssessor).Deposito
    Patrimonio(0).Retirada = Patrimonio(0).Retirada + Patrimonio(intAssessor).Retirada
    Patrimonio(0).Net = Patrimonio(0).Net + Patrimonio(intAssessor).Net
    
    Patrimonio(0).Transf_Entrada = Patrimonio(0).Transf_Entrada + Patrimonio(intAssessor).Transf_Entrada
    Patrimonio(0).Transf_Saida = Patrimonio(0).Transf_Saida + Patrimonio(intAssessor).Transf_Saida
    Patrimonio(0).Transf_Total = Patrimonio(0).Transf_Total + Patrimonio(intAssessor).Transf_Total

    Patrimonio(0).Captacao = Patrimonio(0).Captacao + Patrimonio(intAssessor).Captacao
    
    Patrimonio(0).PL = Patrimonio(0).PL + Patrimonio(intAssessor).PL
    
    intAssessor = intAssessor + 1
    rsAssessor.MoveNext
Wend

'Apaga dados anterior da tabela patrinimo
dbDados.Execute ("Delete * from Patrimonio Where Data=#" & Format(dtDataInicio, "mm/dd/yyyy") & "#")

'Grava dados do banco
Set rsPatrimonio = dbDados.OpenRecordset("Select * from Patrimonio")
For intAssessor = 0 To UBound(Patrimonio)
    rsPatrimonio.AddNew
    rsPatrimonio("Data") = Patrimonio(intAssessor).Data
    rsPatrimonio("Assessor") = Patrimonio(intAssessor).Assessor
    rsPatrimonio("PL") = Patrimonio(intAssessor).PL
    rsPatrimonio("Captacao") = Patrimonio(intAssessor).Captacao
    rsPatrimonio("Deposito") = Patrimonio(intAssessor).Deposito
    rsPatrimonio("Retirada") = Patrimonio(intAssessor).Retirada
    rsPatrimonio("NET") = Patrimonio(intAssessor).Net
    rsPatrimonio("Transf_Entrada") = Patrimonio(intAssessor).Transf_Entrada
    rsPatrimonio("Transf_Saida") = Patrimonio(intAssessor).Transf_Saida
    rsPatrimonio("Transf_Mes") = Patrimonio(intAssessor).Transf_Total
    rsPatrimonio.Update
Next



End Function