<%
' Cria um objeto FileSystemObject
Set fso = Server.CreateObject("Scripting.FileSystemObject")

' Pega o caminho físico do script
path = Server.MapPath(".")

' Chama a função recursiva para listar os arquivos dos subdiretórios
ListFiles path

' Função que recebe um caminho de diretório e lista os arquivos dos subdiretórios
Sub ListFiles(folderPath)

    ' Cria um objeto Folder para o caminho recebido
    Set folder = fso.GetFolder(folderPath)

    ' Percorre os subdiretórios do diretório
    For Each subfolder In folder.SubFolders

      ' Mostra os arquivos do subdiretório na tela
      For Each file In subfolder.Files
        ' Mostra o caminho completo do arquivo
        Response.Write file.Path & "<br>"
      Next

      ' Chama a função recursivamente para o subdiretório
      ListFiles subfolder.Path

    Next

End Sub

%>
