'INSTALAR A REFERÊNCIA System.IO.Compression.FileSystem


Public Sub Main()
    '

    Dim Arquivo(0) As String
    Dim CaminhoArquivo As String = ("C:\Users\jefferson.souza\Desktop\Nova pasta (2)\")



    ' obtem nome do arquivo
    Arquivo = GetFiles(CaminhoArquivo, "*.zip")

    Try
        ZipFile.ExtractToDirectory(Arquivo(0), "C:\Users\jefferson.souza\Desktop\Nova pasta (2)\")

    Catch ex As Exception
        MessageBox.Show("Erro : " & ex.Message, "Erro Arquivo ZIP", MessageBoxButtons.OK, MessageBoxIcon.Error)
    End Try

    Dts.TaskResult = ScriptResults.Success
End Sub