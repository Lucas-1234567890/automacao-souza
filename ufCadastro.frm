VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ufCadastro 
   Caption         =   "Registro de Materiais"
   ClientHeight    =   6330
   ClientLeft      =   -990
   ClientTop       =   -5190
   ClientWidth     =   9615.001
   OleObjectBlob   =   "ufCadastro.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ufCadastro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()

    ' Verifica se o campo "Geradores" foi preenchido
    If cbGeradores.Value = "" Then
        MsgBox "Por favor, selecione um gerador.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Data" foi preenchido e se é uma data válida
    If Not IsDate(txtData.Value) Then
        MsgBox "Por favor, insira uma data válida.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Local" foi preenchido
    If cbTecnico.Value = "" Then
        MsgBox "Por favor, insira o nome do técnico.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Substituição" foi preenchido
    If cbMateriais.Value = "" Then
        MsgBox "Por favor, selecione os materiais.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Substituído por" foi preenchido quando "Substituição" é sim
    If txtInterno.Value = "" Then
        MsgBox "Por favor, insira o número do ID Interno.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Manutenção/Desmobilização" foi preenchido
    If txtExterno.Value = "" Then
        MsgBox "Por favor, selecione o ID Externo.", vbExclamation
        Exit Sub
    End If
    
    ' Verifica se o campo "Quantidade" foi preenchido
    
    If txtQuantidade.Value = "" Then
    MsgBox "Por favor, preenhcer o campo quantidade.", vbExclamation
    Exit Sub
    End If
    

    'linha = thisworbook.Sheets("Cadastro").Range("a1").End(xlDown).Row + 1
    linha = ThisWorkbook.Sheets("Cadastro de materiais").Range("f1000000").End(xlUp).Row + 1
 
    ' Calcula o próximo ID TABELA
    nextID = ThisWorkbook.Sheets("Cadastro de materiais").Cells(linha - 1, 13).Value + 1
 
  Dim valorOriginal As String
  Dim partes() As String
  Dim finalFormatado As String

  valorOriginal = cbGeradores.Value
  partes = Split(valorOriginal, " ") ' separa "55KVA" e "GG11"

  If UBound(partes) = 1 Then
    Dim kva As String, gg As String, numeroGG As String
    kva = Replace(partes(0), "KVA", "")       ' 55
    gg = partes(1)                            ' GG11
    numeroGG = Replace(gg, "GG", "")          ' 11
    finalFormatado = "GERADOR GG-" & kva & numeroGG
  Else
    finalFormatado = cbGeradores.Value ' fallback, caso algo dê errado
  End If

  Sheets("Cadastro de materiais").Cells(linha, 6).Value = finalFormatado

  Sheets("Cadastro de materiais").Cells(linha, 7).Value = DateValue(txtData.Value)
  Sheets("Cadastro de materiais").Cells(linha, 7).NumberFormat = "dd/mm/yyyy"
  Sheets("Cadastro de materiais").Cells(linha, 8).Value = cbTecnico.Value
  Sheets("Cadastro de materiais").Cells(linha, 9).Value = cbMateriais.Value
  Sheets("Cadastro de materiais").Cells(linha, 10).Value = txtInterno.Value
  Sheets("Cadastro de materiais").Cells(linha, 11).Value = txtExterno.Value
  ThisWorkbook.Sheets("Cadastro de materiais").Cells(linha, 13).Value = nextID
  Sheets("Cadastro de materiais").Cells(linha, 12).Value = txtQuantidade.Value

  
  
  
    
  cbMateriais.Value = ""
  txtInterno.Value = ""
  txtExterno.Value = ""
  txtQuantidade.Value = ""
    
End Sub

Private Sub CommandButton2_Click()

   ActiveWorkbook.Worksheets("Cadastro de materiais").ListObjects( _
        "registroMaterial").Sort.SortFields.Clear
    ActiveWorkbook.Worksheets("Cadastro de materiais").ListObjects( _
        "registroMaterial").Sort.SortFields.Add2 Key:=Range( _
        "registroMaterial[[#All],[ID tabela]]"), SortOn:=xlSortOnValues, Order:= _
        xlDescending, DataOption:=xlSortNormal
    With ActiveWorkbook.Worksheets("Cadastro de materiais").ListObjects( _
        "registroMaterial").Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
   Unload ufCadastro
   
End Sub



Private Sub CommandButton3_Click()
ExcluirID.Show
End Sub



Private Sub UserForm_Initialize()
    UltimaLinha = Sheets("opções").Range("c2").End(xlDown).Row
    ultimalinha2 = Sheets("Materiais").Range("g100000").End(xlUp).Row
    txtData.Value = Format(Now, "dd/mm/yyyy")

 
 cbGeradores.RowSource = "opções!c2:c" & UltimaLinha
    
    'essa parte adiciona a lista de vendedroes na caixa de vendedor
    
     UltimaLinha = Sheets("Opções").Range("e2").End(xlDown).Row
    
    cbTecnico.RowSource = "Opções!e2:e" & ultimalinha2
    cbMateriais.RowSource = "Materiais!h5:h" & ultimalinha2
    
     Me.StartUpPosition = 0 ' Posição personalizada
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
    
End Sub
Private Sub cbMateriais_Change()
    ' Define a planilha que contém os materiais
    Set wsMateriais = Sheets("Materiais")
    
    ' Encontra a célula correspondente ao material selecionado na coluna H
    
    Set materialCell = wsMateriais.Range("H5:H" & wsMateriais.Cells(wsMateriais.Rows.Count, "H").End(xlUp).Row).Find(What:=cbMateriais.Value, LookIn:=xlValues, LookAt:=xlWhole)
    
    
    ' Se o material for encontrado
    If Not materialCell Is Nothing Then
        ' Preenche os campos com base na linha do material encontrado
        txtInterno.Value = wsMateriais.Cells(materialCell.Row, "L").Value
        txtExterno.Value = wsMateriais.Cells(materialCell.Row, "M").Value
       
    Else
        ' Se o material não for encontrado, limpa os campos
        txtInterno.Value = ""
        txtExterno.Value = ""
    End If
End Sub




