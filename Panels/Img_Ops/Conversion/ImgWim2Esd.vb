﻿Imports System.Windows.Forms
Imports Microsoft.Dism
Imports System.IO
Imports System.Threading

Public Class ImgWim2Esd

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        If Not ProgressPanel.IsDisposed Then ProgressPanel.Dispose()
        ProgressPanel.imgSrcFile = TextBox1.Text
        ProgressPanel.imgConversionIndex = NumericUpDown1.Value
        ProgressPanel.imgDestFile = TextBox2.Text
        If ComboBox1.SelectedIndex = 0 Then
            ProgressPanel.imgConversionMode = 1
        ElseIf ComboBox1.SelectedIndex = 1 Then
            ProgressPanel.imgConversionMode = 0
        End If
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        ProgressPanel.OperationNum = 991
        Visible = False
        ProgressPanel.ShowDialog(MainForm)
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        TextBox2.Text = SaveFileDialog1.FileName
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        TextBox1.Text = OpenFileDialog1.FileName
        If OpenFileDialog1.FileName.EndsWith("wim", StringComparison.OrdinalIgnoreCase) Then
            ComboBox1.SelectedIndex = 1
        ElseIf OpenFileDialog1.FileName.EndsWith("esd", StringComparison.OrdinalIgnoreCase) Then
            ComboBox1.SelectedIndex = 0
        End If
    End Sub

    Private Sub ImgWim2Esd_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Select Case MainForm.Language
            Case 0
                Select Case My.Computer.Info.InstalledUICulture.ThreeLetterWindowsLanguageName
                    Case "ENU", "ENG"
                        Text = "Convert image"
                        Label1.Text = Text
                        Label2.Text = "Source image file:"
                        Label3.Text = "Format of converted image:"
                        Label5.Text = "Destination image file:"
                        Label7.Text = "Index:"
                        LinkLabel1.Text = "Which format do I choose?"
                        Button1.Text = "Browse..."
                        Button2.Text = "Browse..."
                        OK_Button.Text = "OK"
                        Cancel_Button.Text = "Cancel"
                        ListView1.Columns(0).Text = "Index"
                        ListView1.Columns(1).Text = "Image name"
                        ListView1.Columns(2).Text = "Image description"
                        ListView1.Columns(3).Text = "Image version"
                        GroupBox1.Text = "Source"
                        GroupBox2.Text = "Options"
                        GroupBox3.Text = "Destination"
                        OpenFileDialog1.Title = "Specify the source image file you want to convert"
                        SaveFileDialog1.Title = "Where will the target image be stored?"
                    Case "ESN"
                        Text = "Convertir imagen"
                        Label1.Text = Text
                        Label2.Text = "Archivo de imagen de origen:"
                        Label3.Text = "Formato de imagen convertida:"
                        Label5.Text = "Archivo de imagen de destino:"
                        Label7.Text = "Índice:"
                        LinkLabel1.Text = "¿Qué formato escojo?"
                        Button1.Text = "Examinar..."
                        Button2.Text = "Examinar..."
                        OK_Button.Text = "Aceptar"
                        Cancel_Button.Text = "Cancelar"
                        ListView1.Columns(0).Text = "Índice"
                        ListView1.Columns(1).Text = "Nombre de imagen"
                        ListView1.Columns(2).Text = "Descripción de imagen"
                        ListView1.Columns(3).Text = "Versión de imagen"
                        GroupBox1.Text = "Origen"
                        GroupBox2.Text = "Opciones"
                        GroupBox3.Text = "Destino"
                        OpenFileDialog1.Title = "Especifique el archivo de imagen de origen que desea convertir"
                        SaveFileDialog1.Title = "¿Dónde se almacenará el archivo de imagen de destino?"
                    Case "FRA"
                        Text = "Convertir l'image"
                        Label1.Text = Text
                        Label2.Text = "Fichier de l'image source :"
                        Label3.Text = "Format de l'image convertie :"
                        Label5.Text = "Fichier de l'image de destination :"
                        Label7.Text = "Index :"
                        LinkLabel1.Text = "Quel format dois-je choisir ?"
                        Button1.Text = "Parcourir..."
                        Button2.Text = "Parcourir..."
                        OK_Button.Text = "OK"
                        Cancel_Button.Text = "Annuler"
                        ListView1.Columns(0).Text = "Index"
                        ListView1.Columns(1).Text = "Nom de l'image"
                        ListView1.Columns(2).Text = "Description de l'image"
                        ListView1.Columns(3).Text = "Version de l'image"
                        GroupBox1.Text = "Source"
                        GroupBox2.Text = "Paramètres"
                        GroupBox3.Text = "Destination"
                        OpenFileDialog1.Title = "Spécifiez le fichier de l'image source que vous souhaitez convertir"
                        SaveFileDialog1.Title = "Où l'image cible sera-t-elle stockée ?"
                    Case "PTB", "PTG"
                        Text = "Converter imagem"
                        Label1.Text = Text
                        Label2.Text = "Ficheiro de imagem de origem:"
                        Label3.Text = "Formato da imagem convertida:"
                        Label5.Text = "Ficheiro de imagem de destino:"
                        Label7.Text = "Índice:"
                        LinkLabel1.Text = "Qual o formato que devo escolher?"
                        Button1.Text = "Navegar..."
                        Button2.Text = "Navegar..."
                        OK_Button.Text = "OK"
                        Cancel_Button.Text = "Cancelar"
                        ListView1.Columns(0).Text = "Índice"
                        ListView1.Columns(1).Text = "Nome da imagem"
                        ListView1.Columns(2).Text = "Descrição da imagem"
                        ListView1.Columns(3).Text = "Versão da imagem"
                        GroupBox1.Text = "Fonte"
                        GroupBox2.Text = "Opções"
                        GroupBox3.Text = "Destino"
                        OpenFileDialog1.Title = "Especifique o ficheiro de imagem de origem que pretende converter"
                        SaveFileDialog1.Title = "Onde será guardada a imagem de destino?"
                    Case "ITA"
                        Text = "Convertire immagine"
                        Label1.Text = Text
                        Label2.Text = "File immagine di origine:"
                        Label3.Text = "Formato dell'immagine convertita:"
                        Label5.Text = "File immagine di destinazione:"
                        Label7.Text = "Indice:"
                        LinkLabel1.Text = "Quale formato devo scegliere?"
                        Button1.Text = "Sfoglia..."
                        Button2.Text = "Sfoglia..."
                        OK_Button.Text = "OK"
                        Cancel_Button.Text = "Annullare"
                        ListView1.Columns(0).Text = "Indice"
                        ListView1.Columns(1).Text = "Nome dell'immagine"
                        ListView1.Columns(2).Text = "Descrizione dell'immagine"
                        ListView1.Columns(3).Text = "Versione dell'immagine"
                        GroupBox1.Text = "Sorgente"
                        GroupBox2.Text = "Opzioni"
                        GroupBox3.Text = "Destinazione"
                        OpenFileDialog1.Title = "Specificare il file immagine di origine che si desidera convertire"
                        SaveFileDialog1.Title = "Dove verrà memorizzata l'immagine di destinazione?"
                End Select
            Case 1
                Text = "Convert image"
                Label1.Text = Text
                Label2.Text = "Source image file:"
                Label3.Text = "Format of converted image:"
                Label5.Text = "Destination image file:"
                Label7.Text = "Index:"
                LinkLabel1.Text = "Which format do I choose?"
                Button1.Text = "Browse..."
                Button2.Text = "Browse..."
                OK_Button.Text = "OK"
                Cancel_Button.Text = "Cancel"
                ListView1.Columns(0).Text = "Index"
                ListView1.Columns(1).Text = "Image name"
                ListView1.Columns(2).Text = "Image description"
                ListView1.Columns(3).Text = "Image version"
                GroupBox1.Text = "Source"
                GroupBox2.Text = "Options"
                GroupBox3.Text = "Destination"
                OpenFileDialog1.Title = "Specify the source image file you want to convert"
                SaveFileDialog1.Title = "Where will the target image be stored?"
            Case 2
                Text = "Convertir imagen"
                Label1.Text = Text
                Label2.Text = "Archivo de imagen de origen:"
                Label3.Text = "Formato de imagen convertida:"
                Label5.Text = "Archivo de imagen de destino:"
                Label7.Text = "Índice:"
                LinkLabel1.Text = "¿Qué formato escojo?"
                Button1.Text = "Examinar..."
                Button2.Text = "Examinar..."
                OK_Button.Text = "Aceptar"
                Cancel_Button.Text = "Cancelar"
                ListView1.Columns(0).Text = "Índice"
                ListView1.Columns(1).Text = "Nombre de imagen"
                ListView1.Columns(2).Text = "Descripción de imagen"
                ListView1.Columns(3).Text = "Versión de imagen"
                GroupBox1.Text = "Origen"
                GroupBox2.Text = "Opciones"
                GroupBox3.Text = "Destino"
                OpenFileDialog1.Title = "Especifique el archivo de imagen de origen que desea convertir"
                SaveFileDialog1.Title = "¿Dónde se almacenará el archivo de imagen de destino?"
            Case 3
                Text = "Convertir l'image"
                Label1.Text = Text
                Label2.Text = "Fichier de l'image source :"
                Label3.Text = "Format de l'image convertie :"
                Label5.Text = "Fichier de l'image de destination :"
                Label7.Text = "Index :"
                LinkLabel1.Text = "Quel format dois-je choisir ?"
                Button1.Text = "Parcourir..."
                Button2.Text = "Parcourir..."
                OK_Button.Text = "OK"
                Cancel_Button.Text = "Annuler"
                ListView1.Columns(0).Text = "Index"
                ListView1.Columns(1).Text = "Nom de l'image"
                ListView1.Columns(2).Text = "Description de l'image"
                ListView1.Columns(3).Text = "Version de l'image"
                GroupBox1.Text = "Source"
                GroupBox2.Text = "Paramètres"
                GroupBox3.Text = "Destination"
                OpenFileDialog1.Title = "Spécifiez le fichier de l'image source que vous souhaitez convertir"
                SaveFileDialog1.Title = "Où l'image cible sera-t-elle stockée ?"
            Case 4
                Text = "Converter imagem"
                Label1.Text = Text
                Label2.Text = "Ficheiro de imagem de origem:"
                Label3.Text = "Formato da imagem convertida:"
                Label5.Text = "Ficheiro de imagem de destino:"
                Label7.Text = "Índice:"
                LinkLabel1.Text = "Qual o formato que devo escolher?"
                Button1.Text = "Navegar..."
                Button2.Text = "Navegar..."
                OK_Button.Text = "OK"
                Cancel_Button.Text = "Cancelar"
                ListView1.Columns(0).Text = "Índice"
                ListView1.Columns(1).Text = "Nome da imagem"
                ListView1.Columns(2).Text = "Descrição da imagem"
                ListView1.Columns(3).Text = "Versão da imagem"
                GroupBox1.Text = "Fonte"
                GroupBox2.Text = "Opções"
                GroupBox3.Text = "Destino"
                OpenFileDialog1.Title = "Especifique o ficheiro de imagem de origem que pretende converter"
                SaveFileDialog1.Title = "Onde será guardada a imagem de destino?"
            Case 5
                Text = "Convertire immagine"
                Label1.Text = Text
                Label2.Text = "File immagine di origine:"
                Label3.Text = "Formato dell'immagine convertita:"
                Label5.Text = "File immagine di destinazione:"
                Label7.Text = "Indice:"
                LinkLabel1.Text = "Quale formato devo scegliere?"
                Button1.Text = "Sfoglia..."
                Button2.Text = "Sfoglia..."
                OK_Button.Text = "OK"
                Cancel_Button.Text = "Annullare"
                ListView1.Columns(0).Text = "Indice"
                ListView1.Columns(1).Text = "Nome dell'immagine"
                ListView1.Columns(2).Text = "Descrizione dell'immagine"
                ListView1.Columns(3).Text = "Versione dell'immagine"
                GroupBox1.Text = "Sorgente"
                GroupBox2.Text = "Opzioni"
                GroupBox3.Text = "Destinazione"
                OpenFileDialog1.Title = "Specificare il file immagine di origine che si desidera convertire"
                SaveFileDialog1.Title = "Dove verrà memorizzata l'immagine di destinazione?"
        End Select
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Win10Title.BackColor = Color.FromArgb(48, 48, 48)
            BackColor = Color.FromArgb(31, 31, 31)
            ForeColor = Color.White
            GroupBox1.ForeColor = Color.White
            GroupBox2.ForeColor = Color.White
            GroupBox3.ForeColor = Color.White
            TextBox1.BackColor = Color.FromArgb(31, 31, 31)
            TextBox2.BackColor = Color.FromArgb(31, 31, 31)
            ComboBox1.BackColor = Color.FromArgb(31, 31, 31)
            NumericUpDown1.BackColor = Color.FromArgb(31, 31, 31)
            ListView1.BackColor = Color.FromArgb(31, 31, 31)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Win10Title.BackColor = Color.White
            BackColor = Color.FromArgb(238, 238, 242)
            ForeColor = Color.Black
            GroupBox1.ForeColor = Color.Black
            GroupBox2.ForeColor = Color.Black
            GroupBox3.ForeColor = Color.Black
            TextBox1.BackColor = Color.FromArgb(238, 238, 242)
            TextBox2.BackColor = Color.FromArgb(238, 238, 242)
            ComboBox1.BackColor = Color.FromArgb(238, 238, 242)
            NumericUpDown1.BackColor = Color.FromArgb(238, 238, 242)
            ListView1.BackColor = Color.FromArgb(238, 238, 242)
        End If
        TextBox1.ForeColor = ForeColor
        TextBox2.ForeColor = ForeColor
        ComboBox1.ForeColor = ForeColor
        NumericUpDown1.ForeColor = ForeColor
        ListView1.ForeColor = ForeColor
        If Environment.OSVersion.Version.Major = 10 Then
            Text = ""
            Win10Title.Visible = True
        End If
        Dim handle As IntPtr = MainForm.GetWindowHandle(Me)
        If MainForm.IsWindowsVersionOrGreater(10, 0, 18362) Then MainForm.EnableDarkTitleBar(handle, MainForm.BackColor = Color.FromArgb(48, 48, 48))
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveFileDialog1.ShowDialog()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        SaveFileDialog1.Filter = UCase(ComboBox1.SelectedItem) & " files|*." & LCase(ComboBox1.SelectedItem)
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        If TextBox1.Text <> "" And File.Exists(TextBox1.Text) Then
            If MainForm.MountedImageDetectorBW.IsBusy Then
                MainForm.MountedImageDetectorBWRestarterTimer.Enabled = False
                MainForm.MountedImageDetectorBW.CancelAsync()
                While MainForm.MountedImageDetectorBW.IsBusy
                    Application.DoEvents()
                    Thread.Sleep(500)
                End While
            End If
            MainForm.WatcherTimer.Enabled = False
            If MainForm.WatcherBW.IsBusy Then MainForm.WatcherBW.CancelAsync()
            While MainForm.WatcherBW.IsBusy
                Application.DoEvents()
                Thread.Sleep(100)
            End While
            Try
                ListView1.Items.Clear()
                DismApi.Initialize(DismLogLevel.LogErrors)
                Dim imgInfoCollection As DismImageInfoCollection = DismApi.GetImageInfo(TextBox1.Text)
                NumericUpDown1.Maximum = imgInfoCollection.Count
                For Each imgInfo As DismImageInfo In imgInfoCollection
                    ListView1.Items.Add(New ListViewItem(New String() {imgInfo.ImageIndex, imgInfo.ImageName, imgInfo.ImageDescription, imgInfo.ProductVersion.ToString()}))
                Next
            Catch ex As Exception
                MsgBox("Could not get index information for this image file", vbOKOnly + vbCritical, Label1.Text)
            Finally
                Try
                    DismApi.Shutdown()
                Catch ex As Exception

                End Try
            End Try
        End If
    End Sub
End Class
