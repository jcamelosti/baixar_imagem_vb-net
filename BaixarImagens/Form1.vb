Imports System.IO
Imports System.Net

Public Class Form1
    Private Sub ParseCSV(ByVal path As String, ByVal localSalvar As String)
        Dim numNota As Object
        Dim marca As Object
        Dim linkImagem As Object

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(path)

            MyReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.Delimited
            MyReader.Delimiters = New String() {";"}

            MyReader.ReadLine() ' skip header
            'Loop through all of the fields in the file. 
            'If any lines are corrupt, report an error and continue parsing. 
            While Not MyReader.EndOfData
                Try
                    'currentRow = MyReader.ReadFields()
                    Dim fields = MyReader.ReadFields()

                    numNota = fields(1).Replace("/", "-")
                    marca = fields(3)
                    linkImagem = fields(14)

                    ''Dim a As String
                    ''a = "Index = " & numNota & "oxygenSaturation = " & marca & "pulse = " & linkImagem
                    ''MessageBox.Show(a, "Has been Changed", MessageBoxButtons.OKCancel)

                    ' Include code here to handle the row.

                    DownloadImagem(linkImagem, marca, numNota, localSalvar)
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                " is invalid.  Skipping")
                End Try
            End While
        End Using
    End Sub

    Private Sub DownloadImagem(ByVal url As String, ByVal nomePasta As String, ByVal notaNome As String, ByVal localSalvar As String)
        Using client As New WebClient
            ' Set one of the headers.
            client.Headers("User-Agent") = "Mozilla/4.0"
            Try
                Directory.CreateDirectory(localSalvar + "\" + nomePasta)
            Catch
                MessageBox.Show("Não foi possível criar o diretório")
            End Try
            ' Download data as byte array.
            client.DownloadFile(url, localSalvar + "\" + nomePasta + "\" + notaNome + ".jpg")
        End Using
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim fd As OpenFileDialog = New OpenFileDialog()
        Dim strFileName As String
        Dim fb As FolderBrowserDialog = New FolderBrowserDialog()

        fd.Title = "Selecionar o Arquivo CSV"
        fd.InitialDirectory = "C:\"
        fd.Filter = "All files (*.*)|*.*|All files (*.*)|*.*"
        fd.FilterIndex = 2
        fd.RestoreDirectory = True

        If fd.ShowDialog() = DialogResult.OK Then
            strFileName = fd.FileName

            fb.Description = "Selecione uma pasta para realizar o Download das Imagens"
            fb.RootFolder = Environment.SpecialFolder.MyComputer
            fb.ShowNewFolderButton = True

            If fb.ShowDialog = Windows.Forms.DialogResult.OK Then
                ParseCSV(strFileName, fb.SelectedPath)
            End If
        End If
    End Sub
End Class
