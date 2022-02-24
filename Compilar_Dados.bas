Attribute VB_Name = "Compilar_Dados"
Option Explicit

Dim valor As Double
Dim doc As String
Dim data As Date
Sub AtualizarCompilado()

'LÓGICA DO CODIGO:

'Para a conta 1
    'Entrar na aba Conta 1
    'atualizar valores das variaveis
    'ir para aba consolidação
    'verificar se já foi registrados da mesma conta, tipo e data
        'se existir, acrescentar informações na mesma linha.
        'se não, registrar uma nova linha.
    'ir para linha seguinte da conta 1
    'loop ate fim da conta 1
    
'Repetir o processo para a Conta 2

    Application.Calculation = xlCalculationManual
    Application.ScreenUpdating = False


consolidar ("Conta 1")
consolidar ("Conta 2")



Sheets("Consolidação de Contas").Activate
    
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    
    
End Sub
Sub consolidar(nome_aba As String)

Dim range1, cell As Range

Sheets(nome_aba).Activate

'inicializa a análise da aba da conta.

Set range1 = Range("A1", Range("A1").End(xlDown))

'para cada linha do intervalo de valores da aba da conta,
For Each cell In range1
    If AnalisarLinha(cell) Then
        Call RegistrarLinha
    End If

Next



End Sub
Function AnalisarLinha(ByVal cell As Range) As Boolean 'retorna verdadeiro se a movimentação foi finalizada e falso no caso contrário

    If cell.Offset(0, 5).Value = "Pendente" Then
        AnalisarLinha = False
    ElseIf cell.Offset(0, 5).Value = "Finalizado" Then
        data = cell.Value
        doc = cell.Offset(0, 4).Value
        Call PegarValor(cell)
        AnalisarLinha = True
    Else
        AnalisarLinha = False
    End If
    
    
End Function
Sub PegarValor(cell As Range) 'analisará se a movimentação é entrada ou saída. Se for saída atribuirá o valor como negativo
        
    If cell.Offset(0, 2).Value = "Entrada" Then
        valor = cell.Offset(0, 3).Value
    ElseIf cell.Offset(0, 2).Value = "Saída" Then
        valor = -cell.Offset(0, 3).Value
    Else
        valor = 0
        cell.Offset(0, 8).Value = "Não compilado"
    End If

End Sub
Sub RegistrarLinha() 'registrará na aba de consolidação as informações retiradas das abas de conta
    Dim nome_aba_atual As String
    Dim range_consolidado, cell As Range
    Dim var_check As Boolean
    
    nome_aba_atual = ActiveSheet.name
    Sheets("Consolidação de Contas").Activate
    
    Set range_consolidado = Range(Range("A1"), Range("A1").End(xlDown).Offset(1, 0))
    
    For Each cell In range_consolidado
        If cell.Value = "" Then
        
            cell.Value = data
            cell.Offset(0, 5).Value = doc
            cell.Offset(0, 4).Value = nome_aba_atual
            
            If valor < 0 Then
                cell.Offset(0, 1).Value = "Saída"
                cell.Offset(0, 3).Value = valor
            ElseIf valor > 0 Then
                cell.Offset(0, 1).Value = "Entrada"
                cell.Offset(0, 2).Value = valor
            End If
            
            
            Exit For

        ElseIf cell.Value = data Then '
            var_check = False 'essa variável checa se é o mesmo tipo de movimentação e da mesma conta de origem. Se não for alguma das duas opções, então não será acrescentado na mesma linha e o registro deverá ser feito em uma nova linha
            If cell.Offset(0, 1).Value = "Entrada" And valor > 0 And cell.Offset(0, 4).Value = nome_aba_atual Then
                var_check = True
                cell.Offset(0, 2).Value = cell.Offset(0, 2).Value + valor
            ElseIf cell.Offset(0, 1).Value = "Saída" And valor < 0 And cell.Offset(0, 4).Value = nome_aba_atual Then
                var_check = True
                cell.Offset(0, 3).Value = cell.Offset(0, 3).Value + valor
                
            End If
            
            If var_check Then
                cell.Offset(0, 5).Value = cell.Offset(0, 5).Value & ";" & doc 'concatena os valores de doc separando-os por ;
                Exit For
            End If
        
        End If
    Next
    
    'retorna para a aba da conta para seguir com as linhas seguintes
    Sheets(nome_aba_atual).Activate
    
End Sub
