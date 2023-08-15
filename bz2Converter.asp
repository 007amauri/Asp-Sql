<%
' Cria um objeto FileSystemObject
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Pega o caminho físico do script
path = Server.MapPath(".")

' Chama a função recursiva para listar os arquivos dos subdiretórios e exibir os arquivos que não têm a extensão .bz2
ListFiles path

' Função que recebe um caminho de diretório e lista os arquivos dos subdiretórios e exibe os arquivos que não têm a extensão .bz2
Sub ListFiles(folderPath)

    ' Cria um objeto Folder para o caminho recebido
    Set folder = fso.GetFolder(folderPath)

    ' Percorre os subdiretórios do diretório
    For Each subfolder In folder.SubFolders

      ' Mostra os arquivos do subdiretório na tela
      For Each file In subfolder.Files

        ' Verifica se o arquivo tem a extensão .bz2
        If Right(file.Name, 4) <> ".bz2" Then
          ' Exibe o arquivo na tela
		  Response.Write("bzip2.exe -z -f """ & file.Path & """<br>")
        End If

      Next

      ' Chama a função recursivamente para o subdiretório
      ListFiles subfolder.Path

    Next

End Sub
Response.Write("Echo ""<font color=violet>Os arquivos foram convertidos em .bz2</font>""<br>")
Response.Write("pause")

%>
