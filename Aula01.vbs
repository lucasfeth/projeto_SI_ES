
    Dim lista, item, numeros, maior, menor, resp
    maior = 0
    menor = 999



    resp = InputBox("[0] For (Inputa quantos numeros vai inserir)" & vbNewLine & _ 
                    "[1] While (A cada volta pergunta se deseja continuar)", _
                vbQuestion + vbYesNo)


    If resp = 0 Then
        call maiormenorcomfor()
    Else
        call maiormenorcomdowhile()
    End If


Sub maiormenorcomfor() 
    
    'criei uma lista e depois fiz analise da lista de forma separada'
    numeros = InputBox("Quantidade de numeros a analisar" & i)

    Set lista = CreateObject("System.Collections.ArrayList")
    for i = 1 to numeros
        item = CInt(InputBox("Numero " & i))
        lista.add item
    next

    For each n in lista
        If n > maior Then
            maior = n
        End If 
        if n < menor Then
            menor = n
        End If
    Next

    Dim mensagem
    mensagem = ("Maior: " & maior & vbNewLine & "Menor: " & menor)

    MsgBox mensagem
End Sub

Sub maiormenorcomdowhile()

    'fui perguntando o numero, em seguida perguntando se o usuario quer inserir outro'
    resp = vbYes
    Set lista = CreateObject("System.Collections.ArrayList")
    Do While resp = vbYes
        item = CInt(InputBox("Numero " & i))
        If item > maior Then
            maior = item
        End If 
        if item < menor Then
            menor = item
        End If
        resp = MsgBox("Você deseja continuar?", vbYesNo + vbQuestion, "Confirmação")
    Loop
    MsgBox("Maior: " & maior & vbNewLine & "Menor: " & menor)
End Sub