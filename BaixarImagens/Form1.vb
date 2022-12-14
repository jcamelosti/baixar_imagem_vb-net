Imports System.IO
Imports System.Net
Imports Microsoft.Win32

Public Class Form1
    Function CarRemove(Texto As String, Caracteres As String) As String
        Dim Ic As Integer, It As Integer, Inicio As Integer, Pos As Integer, Caracter As String

        For Ic = 1 To Len(Caracteres)
            Caracter = Mid(Caracteres, Ic, 1)
            Pos = 1
            Inicio = 1
            If InStr(Inicio, Texto, Caracter) > 0 Then
                For It = 1 To Len(Texto)
                    Pos = InStr(Inicio, Texto, Caracter)
                    CarRemove = Mid(Texto, 1, Pos - 1) & Mid(Texto, Pos + 1)
                    Inicio = Pos
                Next It
#Disable Warning BC42104 ' A variável é usada antes de receber um valor
                Texto = CarRemove
#Enable Warning BC42104 ' A variável é usada antes de receber um valor
                Ic = Ic - 1
            Else
                CarRemove = Texto
            End If
        Next Ic
#Disable Warning BC42105 ' A função não retorna um valor em todos os caminhos de código
    End Function
#Enable Warning BC42105 ' A função não retorna um valor em todos os caminhos de código

    Private Sub ParseCSV(ByVal path As String, ByVal localSalvar As String)
        Dim numNota As Object
        Dim marca As Object
        Dim linkImagem As Object

        Using MyReader As New Microsoft.VisualBasic.FileIO.TextFieldParser(path)

            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.Delimiters = New String() {";"}

            MyReader.ReadLine() ' skip header
            'Loop through all of the fields in the file. 
            'If any lines are corrupt, report an error and continue parsing. 
            While Not MyReader.EndOfData
                Try
                    'currentRow = MyReader.ReadFields()
                    Dim fields = MyReader.ReadFields()

                    numNota = fields(1).Replace("/", "-")
                    marca = CarRemove(fields(3), "\/:*?'<>|@")
                    linkImagem = fields(14)

                    If Trim(numNota) <> "" And Trim(marca) <> "" And Trim(linkImagem) <> "" Then
                        DownloadImagem(linkImagem, marca, numNota, localSalvar)
                    End If
                Catch ex As Microsoft.VisualBasic.FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                " is invalid.  Skipping")
                End Try
            End While
        End Using

        RichTextBox1.AppendText("Concluído")
    End Sub

    Private Sub QuebrarCSV(ByVal path As String, ByVal localSalvar As String)
        Dim numNota As Object
        Dim marca As Object
        Dim linkImagem As Object
        Dim header As Object
        Dim qtdRealizada As Integer
        Dim countArquivo As Integer
        Dim registros As New ArrayList

        Using MyReader As New FileIO.TextFieldParser(path)

            MyReader.TextFieldType = FileIO.FieldType.Delimited
            MyReader.Delimiters = New String() {TextBox2.Text}

            MyReader.ReadLine() ' skip header
            'Loop through all of the fields in the file. 
            'If any lines are corrupt, report an error and continue parsing. 
            header = "numero_nota" + TextBox2.Text + "marca" + TextBox2.Text + "link"

            qtdRealizada = 0
            countArquivo = 0

            While Not MyReader.EndOfData
                Try
                    'currentRow = MyReader.ReadFields()
                    Dim fields = MyReader.ReadFields()

                    numNota = fields(1).Replace("/", "-")
                    marca = CarRemove(fields(3), "\/:*?'<>|@")
                    linkImagem = fields(14)

                    registros.Add(New Registro(numNota, marca, linkImagem))
                Catch ex As FileIO.MalformedLineException
                    MsgBox("Line " & ex.Message &
                " is invalid.  Skipping")
                End Try

                qtdRealizada += 1

                If qtdRealizada = Int(TextBox1.Text) Then
                    qtdRealizada = 0

                    ''GRAVAR CSV
                    Using sw As StreamWriter = File.CreateText(localSalvar + "\" + countArquivo.ToString + ".csv")
                        sw.WriteLine(header)

                        For Each row As Registro In registros
                            Dim rowData As String

                            rowData = row.NumNota + TextBox2.Text + row.Marca + TextBox2.Text + row.Link

                            sw.WriteLine(rowData)
                        Next
                    End Using

                    registros = New ArrayList
                    countArquivo += 1
                End If
            End While
        End Using

        RichTextBox1.AppendText("Concluído")
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
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

            fb.Description = "Selecione uma pasta para os Arquivos CSV serem Gerados."
            fb.RootFolder = Environment.SpecialFolder.MyComputer
            fb.ShowNewFolderButton = True

            If fb.ShowDialog = Windows.Forms.DialogResult.OK Then
                QuebrarCSV(strFileName, fb.SelectedPath)
            End If
        End If
    End Sub
End Class
