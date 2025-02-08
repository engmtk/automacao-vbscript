Dim http, json, fso, file, inicio, fim, registros, i
Set http = CreateObject("MSXML2.XMLHTTP")
Set fso = CreateObject("Scripting.FileSystemObject")

' URL da API
url = "https://jsonplaceholder.typicode.com/users"

' Fazendo requisição HTTP GET
http.Open "GET", url, False
http.Send

' Criando arquivo TXT
Set file = fso.CreateTextFile("C:\Temp\dados_api.txt", True)

' Verificando se a resposta é válida
If http.Status = 200 Then
    json = http.responseText
    WScript.Echo "JSON RECEBIDO: " & json  ' Debug para ver o JSON

    ' Contando registros no JSON manualmente
    registros = 0
    inicio = 1
    
    Do While InStr(inicio, json, """id"":") > 0
        registros = registros + 1
        inicio = InStr(inicio, json, """id"":") + 4
    Loop

    WScript.Echo "Registros encontrados: " & registros

    ' Iterando sobre os registros manualmente
    inicio = 1
    For i = 1 To registros
        id = ExtrairValor(json, """id"":", inicio)
        name = ExtrairValor(json, """name"":", inicio)
        email = ExtrairValor(json, """email"":", inicio)
        phone = ExtrairValor(json, """phone"":", inicio)
        
        ' Debug: Exibir no console antes de salvar
        WScript.Echo "Registro " & i & " -> ID: " & id & ", Nome: " & name & ", Email: " & email & ", Telefone: " & phone

        ' Escrevendo os valores no TXT
        file.WriteLine "ID: " & id
        file.WriteLine "Nome: " & name
        file.WriteLine "Email: " & email
        file.WriteLine "Telefone: " & phone
        file.WriteLine "------------------------------------"
    Next

    file.Close
    WScript.Echo "Dados salvos com sucesso em C:\Temp\dados_api.txt"
Else
    WScript.Echo "Erro ao acessar a API: " & http.Status
End If

' Função para extrair valores do JSON
Function ExtrairValor(texto, chave, ByRef inicio)
    Dim fim
    ExtrairValor = "N/A"

    inicio = InStr(inicio, texto, chave)
    If inicio > 0 Then
        inicio = inicio + Len(chave)
        If Mid(texto, inicio, 1) = """" Then inicio = inicio + 1 ' Pula aspas se houver
        fim = InStr(inicio, texto, ",")
        If fim = 0 Then fim = InStr(inicio, texto, "}")
        If fim > inicio Then
            ExtrairValor = Trim(Mid(texto, inicio, fim - inicio))
            ' Removendo aspas extras se existirem
            If Left(ExtrairValor, 1) = """" Then
                ExtrairValor = Mid(ExtrairValor, 2, Len(ExtrairValor) - 2)
            End If
        End If
    End If
End Function
