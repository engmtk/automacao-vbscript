Descrição do projeto, estava desenferrujando sobre VB que é extremamente poderoso mas que caiu em desuso em razão das varias tecnologias, mas ainda sim, vc encontra muito por aí e bem feitinho, sem duvida que atende a necessidade, eu mesmo ja implementei varios scripts para automatizar stop start em serviços, expurgos entre outros.
Esse projeto captura e grava os registros num arquivo .txt encontrados em qualquer api para que voce trabalha-los como quiser, tendo os registros em mãos.
É bem simples, mas muito eficaz.

TODA CONTRIBUIÇÃO é bem vinda!

Automação com VBScript - Extração de Dados de API para Arquivo TXT  
Este projeto realiza a **extração de dados de uma API** e os **salva em um arquivo TXT**, utilizando **VBScript (.vbs)** para automação sem dependências externas. O objetivo é demonstrar como consumir APIs e armazenar informações em um ambiente **sem suporte a PowerShell ou Python**.

Funcionalidades
✅ Faz uma requisição HTTP GET para a API pública https://jsonplaceholder.typicode.com/users 

✅ Processa o JSON retornado e extrai **ID, Nome, E-mail e Telefone**.  

✅ Formata os dados e os salva em um **arquivo TXT (`dados_api.txt`)**.  

✅ Exibe os registros extraídos diretamente no console.  

✅ **Independente de PowerShell, Python ou bibliotecas externas** (ideal para ambientes restritos).  
Como Usar:
 1 - **Crie a pasta de automação**

OBS.: os 2 arquivos de configuração o  devem estar na mesma pasta, o local onde vc pretende salvar o arquivo, fica a seu criterio. 
1-api_to_txt.vbs
2-api_to_csv.bat 

LET's GOOO!

Primeiro, crie uma pasta por exemplo C:\Automacao e dentro dela os 2 arquivos de configuração

Copie e cole o codigo num arquivo .txt e salve como api_to_csv.bat.
@echo off
cscript //nologo D:\automacao\api_to_csv.vbs
pause

Crie um segundo arquivo .txt copie e cole o codigo e salve como api_to_txt.vbs

Com isso ele vai criar um arquivo VBScript

Salve o código abaixo como api_to_txt.vbs em C:\Automacao\:

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

Para que o programa rode, basta executar o .bat

A Saída do Arquivo TXT, deve ser desta forma:

Depois da execução, o arquivo C:\Temp\dados_api.txt conterá os registros extraídos:
ID: 1
Nome: Leanne Graham
Email: Sincere@april.biz
Telefone: 1-770-736-8031 x56442
------------------------------------
ID: 2
Nome: Ervin Howell
Email: Shanna@melissa.tv
Telefone: 010-692-6593 x09125
------------------------------------
ID: 3
Nome: Clementine Bauch
Email: Nathan@yesenia.net
Telefone: 1-463-123-4447
------------------------------------

Contribuições são bem-vindas! Se encontrar algum problema ou tiver sugestões, abra uma issue ou faça um fork do projeto. 
✉️ Contato: 
Linkedin https://www.linkedin.com/in/alexandre-jos%C3%A9-dos-santos-aaaa6110b/
Github https://github.com/engmtk

E-mail alexandre.zero11@gmail.com
