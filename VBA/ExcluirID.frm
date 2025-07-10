VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} ExcluirID 
   Caption         =   "Excluir ID"
   ClientHeight    =   3570
   ClientLeft      =   -855
   ClientTop       =   -4530
   ClientWidth     =   6960
   OleObjectBlob   =   "ExcluirID.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "ExcluirID"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Private Sub CommandButton1_Click()

    ' Declara valor_id como Integer para números inteiros pequenos
    Dim valor_id As Integer
    
    ' Tenta converter o valor da caixa de texto para um número inteiro
    On Error GoTo ErrorHandler
    valor_id = CInt(txtID.Value)
    
    ult_linhas = Range("f1000000").End(xlUp).Row

    For linha = 4 To ult_linhas
        If Cells(linha, 13).Value = valor_id Then
            Range(Cells(linha, 6), Cells(linha, 13)).Delete Shift:=xlUp
        End If
    Next

    ' Limpa a caixa de texto após a operação
    txtID.Value = ""
    
    Exit Sub

ErrorHandler:
    MsgBox "Por favor, insira um número inteiro válido.", vbExclamation
    txtID.SetFocus

End Sub

Private Sub CommandButton2_Click()
Unload ExcluirID
End Sub



Private Sub UserForm_Initialize()
 Me.StartUpPosition = 0 ' Posição personalizada
    Me.Left = Application.Left + (0.5 * Application.Width) - (0.5 * Me.Width)
    Me.Top = Application.Top + (0.5 * Application.Height) - (0.5 * Me.Height)
End Sub

