VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "Registro de Compras"
   ClientHeight    =   8685
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5655
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    Dim objeto As Control
    
    'LOGÍCA DO CÓDIGO
    
        'Abrir o formulário
        'Escolher o local de cadastro
        'Coletar as informações inseridas no box de cadastro
        'Registrar as dentro da planilha e suas respectivas colunas
            'Importante: registrar somente quando o CommandButton1_Click for pressionado
        'Limpar formulário
        
        
        Call Cadastrar
         
    'Limpar formulario
        For Each objeto In UserForm1.Controls
            On Error Resume Next
            objeto.Value = ""
            
        
          
        Next
        
    UserForm1.Hide

End Sub
Sub Cadastrar()
    Dim range1 As Range
    'escolher local de cadastro
        If RefEdit1.Value <> "" Then
            Set range1 = Range(RefEdit1.Value)
        Else
            If Range("A2").Value = "" Then
                Set range1 = Range("A2")
            Else
                Set range1 = Range("A1").End(xlDown).Offset(1, 0)
            End If
            
            
        End If
        
    'colocar informação na coluna respectiva
    range1.Value = UserForm1.ComboBox1.Value
    range1.Offset(0, 1).Value = UserForm1.ListBox1.Value
    range1.Offset(0, 2).Value = UserForm1.ToggleButton1.Value
    range1.Offset(0, 3).Value = UserForm1.CheckBox1.Value
    range1.Offset(0, 4).Value = UserForm1.CheckBox2.Value
    range1.Offset(0, 5).Value = UserForm1.CheckBox3.Value
    range1.Offset(0, 6).Value = UserForm1.CheckBox4.Value
    
    If UserForm1.OptionButton1.Value = True Then
        range1.Offset(0, 7).Value = "PRODUTO"
    Else
        range1.Offset(0, 7).Value = "SERVIÇO"
    End If
    
    If UserForm1.OptionButton3.Value = True Then
        range1.Offset(0, 8).Value = UserForm1.OptionButton3.Caption
    ElseIf UserForm1.OptionButton4.Value = True Then
        range1.Offset(0, 8).Value = UserForm1.OptionButton4.Caption
    ElseIf UserForm1.OptionButton5.Value = True Then
         range1.Offset(0, 8).Value = UserForm1.OptionButton5.Caption
    End If
    
    range1.Offset(0, 9).Value = CDbl(UserForm1.TextBox2.Value)
    range1.Offset(0, 9).Style = "Currency"
    range1.Offset(0, 10).Value = UserForm1.TextBox1.Value
    
    
End Sub
Private Sub ToggleButton1_Click()
    UserForm1.Frame1.Visible = ToggleButton1.Value
    
End Sub

Private Sub UserForm_Initialize()

        With UserForm1.ComboBox1
            .AddItem ("Marketing")
            .AddItem ("Operações")
            .AddItem ("Financeiro")
            .AddItem ("Administrativo")
        End With
        
        UserForm1.ToggleButton1.Caption = "Nota Fiscal Emitida?"
        
        UserForm1.Frame1.Caption = "Impostos"
        UserForm1.Frame1.Visible = False
        
        UserForm1.CheckBox1.Caption = "IR"
        UserForm1.CheckBox2.Caption = "PIS"
        UserForm1.CheckBox3.Caption = "COFINS"
        UserForm1.CheckBox4.Caption = "ISS"
        
        UserForm1.OptionButton1.Caption = "Produto"
        UserForm1.OptionButton2.Caption = "Serviço"
        
    'Multi-página
        UserForm1.OptionButton3.Caption = "Antecipado"
        UserForm1.OptionButton4.Caption = "Na entrega"
        UserForm1.OptionButton5.Caption = "D+30 após entrega"
        
        UserForm1.MultiPage1.Pages(0).Caption = "Prazo de Pagamento"
        UserForm1.MultiPage1.Pages(1).Caption = "Descrição"
        UserForm1.MultiPage1.Pages(2).Caption = "Valor"
    
    'Botão
        UserForm1.CommandButton1.Caption = "Registrar"
        
End Sub
