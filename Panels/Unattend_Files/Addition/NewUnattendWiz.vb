﻿Imports System.IO
Imports Microsoft.VisualBasic.ControlChars
Imports System.Text.Encoding
Imports System.Threading
Imports ScintillaNET
Imports DISMTools.Elements
Imports Microsoft.Dism
Imports System.Net

Public Class NewUnattendWiz

    ' Declare initial vars
    Dim IsInExpress As Boolean = True
    Dim CurrentWizardPage As New UnattendedWizardPage()
    Dim VerifyInPages As New List(Of UnattendedWizardPage.Page)

    Dim DotNetRuntimeSupported As Boolean
    Dim PreferSelfContained As Boolean
    Dim UnattendGenReleaseTag As String = "24122"

    ' Regional Settings Page
    Dim ImageLanguages As New List(Of ImageLanguage)
    Dim UserLocales As New List(Of UserLocale)
    Dim KeyboardIdentifiers As New List(Of KeyboardIdentifier)
    Dim GeoIds As New List(Of GeoId)
    Dim RegionalInteractive As Boolean
    Dim SelectedLanguage As New ImageLanguage()
    Dim SelectedLocale As New UserLocale()
    Dim SelectedKeybIdentifier As New KeyboardIdentifier()
    Dim SelectedGeoId As New GeoId()

    ' System Configuration Page
    Dim SelectedArchitecture As New DismProcessorArchitecture()
    Dim Win11Config As New SVSettings()
    Dim PCName As New ComputerName()

    ' Time Zone Panel
    Dim TimeOffsets As New List(Of TimeOffset)
    Dim TimeOffsetInteractive As Boolean = True
    Dim SelectedOffset As New TimeOffset()

    ' Disk Configuration Panel
    Dim DiskConfigurationInteractive As Boolean = True
    Dim SelectedDiskConfiguration As New DiskConfiguration()

    ' Product Key Panel
    Dim GenericChosen As Boolean = True
    Dim GenericKeys As New List(Of ProductKey)
    Dim SelectedKey As New ProductKey()

    ' User Accounts Panel
    Dim UserAccountsInteractive As Boolean = True
    Dim MicrosoftAccountInteractive As Boolean
    Dim UserAccountsList As New List(Of User)
    Dim AutoLogon As New AutoLogonSettings()
    Dim PasswordObfuscate As Boolean
    Dim SelectedExpirationSettings As New PasswordExpirationSettings()
    Dim SelectedLockdownSettings As New AccountLockdownSettings()

    ' Virtual Machine Panel
    Dim VirtualMachineSupported As Boolean
    Dim SelectedVMSettings As New VirtualMachineSettings()

    ' Wireless Networking Panel
    Dim NetworkConfigInteractive As Boolean = True
    Dim NetworkConfigManualSkip As Boolean = False
    Dim SelectedNetworkConfiguration As New WirelessSettings()

    ' System Telemetry Panel
    Dim SystemTelemetryInteractive As Boolean
    Dim SelectedTelemetrySettings As New SystemTelemetry()

    ' Component Panel
    Dim SystemComponents As New List(Of Component)
    Dim FinalComponents As New List(Of Component)

    ' Space for more pages

    ' Default Settings
    Dim DefaultLanguage As New ImageLanguage()
    Dim DefaultLocale As New UserLocale()
    Dim DefaultKeybIdentifier As New KeyboardIdentifier()
    Dim DefaultGeoId As New GeoId()
    Dim DefaultOffset As New TimeOffset()
    Dim DefaultDiskConfiguration As New DiskConfiguration()
    Dim DefaultExpirationSettings As New PasswordExpirationSettings()
    Dim DefaultLockdownSettings As New AccountLockdownSettings()
    Dim DefaultVMSettings As New VirtualMachineSettings()
    Dim DefaultNetworkConfiguration As New WirelessSettings()
    Dim DefaultSystemComponents As New List(Of Component)

    ' Progress info
    Dim ProgressMessage As String = ""

    Dim SaveTarget As String = ""

    ' Editor Mode
    Dim DefaultContents As String


    ''' <summary>
    ''' Initializes the Scintilla editor
    ''' </summary>
    ''' <param name="fntName">The name of the font used in the Scintilla editor</param>
    ''' <param name="fntSize">The size of the font used in the Scintilla editor</param>
    ''' <remarks></remarks>
    Sub InitScintilla(fntName As String, fntSize As Integer)
        ' Initialize Scintilla editor
        Scintilla1.StyleResetDefault()
        Scintilla2.StyleResetDefault()
        ' Use VS's selection color, as I find it the most natural
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Scintilla1.SelectionBackColor = Color.FromArgb(38, 79, 120)
            Scintilla2.SelectionBackColor = Color.FromArgb(38, 79, 120)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Scintilla1.SelectionBackColor = Color.FromArgb(153, 201, 239)
            Scintilla2.SelectionBackColor = Color.FromArgb(153, 201, 239)
        End If
        Scintilla1.Styles(Style.Default).Font = fntName
        Scintilla1.Styles(Style.Default).Size = fntSize
        Scintilla2.Styles(Style.Default).Font = fntName
        Scintilla2.Styles(Style.Default).Size = fntSize

        ' Set background and foreground colors (from Visual Studio)
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Scintilla1.Styles(Style.Default).BackColor = Color.FromArgb(30, 30, 30)
            Scintilla1.Styles(Style.Default).ForeColor = Color.White
            Scintilla1.Styles(Style.LineNumber).BackColor = Color.FromArgb(30, 30, 30)
            Scintilla2.Styles(Style.Default).BackColor = Color.FromArgb(30, 30, 30)
            Scintilla2.Styles(Style.Default).ForeColor = Color.White
            Scintilla2.Styles(Style.LineNumber).BackColor = Color.FromArgb(30, 30, 30)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Scintilla1.Styles(Style.Default).BackColor = Color.White
            Scintilla1.Styles(Style.Default).ForeColor = Color.Black
            Scintilla1.Styles(Style.LineNumber).BackColor = Color.White
            Scintilla2.Styles(Style.Default).BackColor = Color.White
            Scintilla2.Styles(Style.Default).ForeColor = Color.Black
            Scintilla2.Styles(Style.LineNumber).BackColor = Color.White
        End If
        Scintilla1.StyleClearAll()
        Scintilla2.StyleClearAll()

        ' Use Notepad++'s lexer style colors
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Scintilla1.Styles(Style.Xml.XmlStart).ForeColor = Color.FromArgb(127, 159, 127)
            Scintilla1.Styles(Style.Xml.XmlEnd).ForeColor = Color.FromArgb(127, 159, 127)
            Scintilla1.Styles(Style.Xml.Default).ForeColor = Color.FromArgb(220, 220, 204)
            Scintilla1.Styles(Style.Xml.Comment).ForeColor = Color.FromArgb(127, 159, 127)
            Scintilla1.Styles(Style.Xml.Number).ForeColor = Color.FromArgb(140, 208, 211)
            Scintilla1.Styles(Style.Xml.DoubleString).ForeColor = Color.FromArgb(200, 145, 145)
            Scintilla1.Styles(Style.Xml.SingleString).ForeColor = Color.FromArgb(200, 145, 145)
            Scintilla1.Styles(Style.Xml.Tag).ForeColor = Color.FromArgb(227, 206, 171)
            Scintilla1.Styles(Style.Xml.TagEnd).ForeColor = Color.FromArgb(227, 206, 171)
            Scintilla1.Styles(Style.Xml.TagUnknown).ForeColor = Color.FromArgb(237, 214, 237)
            Scintilla1.Styles(Style.Xml.Attribute).ForeColor = Color.FromArgb(190, 200, 158)
            Scintilla1.Styles(Style.Xml.AttributeUnknown).ForeColor = Color.FromArgb(223, 223, 223)
            Scintilla1.Styles(Style.Xml.CData).ForeColor = Color.FromArgb(200, 145, 145)
            Scintilla1.Styles(Style.Xml.Entity).ForeColor = Color.FromArgb(207, 191, 175)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Scintilla1.Styles(Style.Xml.XmlStart).ForeColor = Color.Red
            Scintilla1.Styles(Style.Xml.XmlEnd).ForeColor = Color.Red
            Scintilla1.Styles(Style.Xml.Default).ForeColor = Color.Black
            Scintilla1.Styles(Style.Xml.Comment).ForeColor = Color.FromArgb(0, 128, 0)
            Scintilla1.Styles(Style.Xml.Number).ForeColor = Color.Red
            Scintilla1.Styles(Style.Xml.DoubleString).ForeColor = Color.FromArgb(128, 0, 255)
            Scintilla1.Styles(Style.Xml.SingleString).ForeColor = Color.FromArgb(128, 0, 255)
            Scintilla1.Styles(Style.Xml.Tag).ForeColor = Color.Blue
            Scintilla1.Styles(Style.Xml.TagEnd).ForeColor = Color.Blue
            Scintilla1.Styles(Style.Xml.TagUnknown).ForeColor = Color.Blue
            Scintilla1.Styles(Style.Xml.Attribute).ForeColor = Color.Red
            Scintilla1.Styles(Style.Xml.AttributeUnknown).ForeColor = Color.Red
            Scintilla1.Styles(Style.Xml.CData).ForeColor = Color.FromArgb(255, 128, 0)
            Scintilla1.Styles(Style.Xml.Entity).ForeColor = Color.Black
        End If
        ' Set lexer
        Scintilla1.LexerName = "xml"

        ' Set line number margin properties
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            Scintilla1.Styles(Style.LineNumber).BackColor = Color.FromArgb(30, 30, 30)
            Scintilla2.Styles(Style.LineNumber).BackColor = Color.FromArgb(30, 30, 30)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            Scintilla1.Styles(Style.LineNumber).BackColor = Color.White
            Scintilla2.Styles(Style.LineNumber).BackColor = Color.White
        End If
        Scintilla1.Styles(Style.LineNumber).ForeColor = Color.FromArgb(165, 165, 165)
        Scintilla2.Styles(Style.LineNumber).ForeColor = Color.FromArgb(165, 165, 165)
        Dim Margin = Scintilla1.Margins(1)
        Margin.Width = 30
        Margin.Type = MarginType.Number
        Margin.Sensitive = True
        Margin.Mask = 0
        Margin = Scintilla2.Margins(1)
        Margin.Width = 30
        Margin.Type = MarginType.Number
        Margin.Sensitive = True
        Margin.Mask = 0

        ' Initialize code folding
        Scintilla1.SetFoldMarginColor(True, Scintilla1.Styles(Style.Default).BackColor)
        Scintilla1.SetFoldMarginColor(True, Scintilla1.Styles(Style.Default).BackColor)
        Scintilla1.SetProperty("fold", "1")
        Scintilla1.SetProperty("fold.compact", "1")
        Scintilla2.SetFoldMarginColor(True, Scintilla1.Styles(Style.Default).BackColor)
        Scintilla2.SetFoldMarginColor(True, Scintilla1.Styles(Style.Default).BackColor)
        Scintilla2.SetProperty("fold", "1")
        Scintilla2.SetProperty("fold.compact", "1")

        ' Configure bookmark margins
        Dim Bookmarks = Scintilla1.Margins(2)
        Bookmarks.Width = 20
        Bookmarks.Sensitive = True
        Bookmarks.Type = MarginType.Symbol
        Bookmarks.Mask = (1 << 2)
        Dim Marker = Scintilla1.Markers(2)
        Marker.Symbol = MarkerSymbol.Circle
        Marker.SetBackColor(Color.FromArgb(255, 0, 59))
        Marker.SetForeColor(Color.Black)
        Marker.SetAlpha(100)
        Bookmarks = Scintilla2.Margins(2)
        Bookmarks.Width = 20
        Bookmarks.Sensitive = True
        Bookmarks.Type = MarginType.Symbol
        Bookmarks.Mask = (1 << 2)
        Marker = Scintilla2.Markers(2)
        Marker.Symbol = MarkerSymbol.Circle
        Marker.SetBackColor(Color.FromArgb(255, 0, 59))
        Marker.SetForeColor(Color.Black)
        Marker.SetAlpha(100)

        ' Set editor caret settings
        Scintilla1.CaretForeColor = ForeColor
        Scintilla2.CaretForeColor = ForeColor


        ' Configure code folding margins
        Scintilla1.Margins(3).Type = MarginType.Symbol
        Scintilla1.Margins(3).Mask = Marker.MaskFolders
        Scintilla1.Margins(3).Sensitive = True
        Scintilla1.Margins(3).Width = 1
        Scintilla2.Margins(3).Type = MarginType.Symbol
        Scintilla2.Margins(3).Mask = Marker.MaskFolders
        Scintilla2.Margins(3).Sensitive = True
        Scintilla2.Margins(3).Width = 1

        ' Set colors for all folding markers
        For x = 25 To 31
            Scintilla1.Markers(x).SetForeColor(Scintilla1.Styles(Style.Default).BackColor)
            Scintilla1.Markers(x).SetBackColor(Scintilla1.Styles(Style.Default).ForeColor)
            Scintilla2.Markers(x).SetForeColor(Scintilla1.Styles(Style.Default).BackColor)
            Scintilla2.Markers(x).SetBackColor(Scintilla1.Styles(Style.Default).ForeColor)
        Next

        ' Folding marker configuration
        Scintilla1.Markers(Marker.Folder).Symbol = MarkerSymbol.BoxPlus
        Scintilla1.Markers(Marker.FolderOpen).Symbol = MarkerSymbol.BoxMinus
        Scintilla1.Markers(Marker.FolderEnd).Symbol = MarkerSymbol.BoxPlusConnected
        Scintilla1.Markers(Marker.FolderMidTail).Symbol = MarkerSymbol.TCorner
        Scintilla1.Markers(Marker.FolderOpenMid).Symbol = MarkerSymbol.BoxMinusConnected
        Scintilla1.Markers(Marker.FolderSub).Symbol = MarkerSymbol.VLine
        Scintilla1.Markers(Marker.FolderTail).Symbol = MarkerSymbol.LCorner
        Scintilla2.Markers(Marker.Folder).Symbol = MarkerSymbol.BoxPlus
        Scintilla2.Markers(Marker.FolderOpen).Symbol = MarkerSymbol.BoxMinus
        Scintilla2.Markers(Marker.FolderEnd).Symbol = MarkerSymbol.BoxPlusConnected
        Scintilla2.Markers(Marker.FolderMidTail).Symbol = MarkerSymbol.TCorner
        Scintilla2.Markers(Marker.FolderOpenMid).Symbol = MarkerSymbol.BoxMinusConnected
        Scintilla2.Markers(Marker.FolderSub).Symbol = MarkerSymbol.VLine
        Scintilla2.Markers(Marker.FolderTail).Symbol = MarkerSymbol.LCorner

        ' Enable folding
        Scintilla1.AutomaticFold = (AutomaticFold.Show Or AutomaticFold.Click Or AutomaticFold.Show)
        Scintilla2.AutomaticFold = (AutomaticFold.Show Or AutomaticFold.Click Or AutomaticFold.Show)
    End Sub

    Function NewKeyVar(key As String) As ProductKey
        Dim pKey As New ProductKey()
        pKey.Valid = True
        pKey.Key = key
        Return pKey
    End Function

    Sub SetDefaultSettings()
        DefaultLanguage.Id = "en-US"
        DefaultLanguage.DisplayName = "English"
        DefaultLocale.Id = "en-US"
        DefaultLocale.DisplayName = "English (United States)"
        DefaultLocale.LCID = "0409"
        DefaultLocale.KeybId = "00000409"
        DefaultLocale.GeoLoc = "244"
        DefaultKeybIdentifier.Id = "00000409"
        DefaultKeybIdentifier.DisplayName = "US"
        DefaultKeybIdentifier.Type = "Keyboard"
        DefaultGeoId.Id = "244"
        DefaultGeoId.DisplayName = "United States"
        DefaultOffset.Id = "UTC"
        DefaultOffset.DisplayName = "(UTC) Coordinated Universal Time"
        DefaultDiskConfiguration.DiskConfigMode = DiskConfigurationMode.AutoDisk0
        DefaultDiskConfiguration.PartStyle = PartitionStyle.GPT
        DefaultDiskConfiguration.ESPSize = 300
        DefaultDiskConfiguration.InstallRecEnv = True
        DefaultDiskConfiguration.RecEnvPartition = RecoveryEnvironmentLocation.WinREPartition
        DefaultDiskConfiguration.RecEnvSize = 1000
        DefaultDiskConfiguration.DiskPartScriptConfig.ScriptContents = ""
        DefaultDiskConfiguration.DiskPartScriptConfig.AutomaticInstall = True
        DefaultDiskConfiguration.DiskPartScriptConfig.TargetDisk.DiskNum = 0
        DefaultDiskConfiguration.DiskPartScriptConfig.TargetDisk.PartNum = 3

        GenericKeys.Add(NewKeyVar("YNMGQ-8RYV3-4PGQ3-C8XTP-7CFBY"))     ' Education
        GenericKeys.Add(NewKeyVar("84NGF-MHBT6-FXBX8-QWJK7-DRR8H"))     ' Education N
        GenericKeys.Add(NewKeyVar("YTMG3-N6DKC-DKB77-7M9GH-8HVX7"))     ' Home
        GenericKeys.Add(NewKeyVar("4CPRK-NM3K3-X6XXQ-RXX86-WXCHW"))     ' Home N
        GenericKeys.Add(NewKeyVar("BT79Q-G7N6G-PGBYW-4YWX6-6F4BT"))     ' Home Simple Language
        GenericKeys.Add(NewKeyVar("VK7JG-NPHTM-C97JM-9MPGT-3V66T"))     ' Pro
        GenericKeys.Add(NewKeyVar("8PTT6-RNW4C-6V7J2-C2D3X-MHBPB"))     ' Pro Education
        GenericKeys.Add(NewKeyVar("GJTYN-HDMQY-FRR76-HVGC7-QPF8P"))     ' Pro Education N
        GenericKeys.Add(NewKeyVar("DXG7C-N36C4-C4HTG-X4T3X-2YV77"))     ' Pro for Workstations
        GenericKeys.Add(NewKeyVar("2B87N-8KFHP-DKV6R-Y2C8J-PKCKT"))     ' Pro N
        GenericKeys.Add(NewKeyVar("WYPNQ-8C467-V2W6J-TX4WX-WT2RQ"))     ' Pro N for Workstations
        GenericKeys.Add(NewKeyVar("XGVPP-NMH47-7TTHJ-W3FW7-8HV2C"))     ' Enterprise

        UserAccountsList.Add(New User(True, "Admin", "", UserGroup.Administrators))
        For i = 1 To 4
            UserAccountsList.Add(New User(False, "", "", UserGroup.Users))
        Next

        DefaultExpirationSettings.Mode = PasswordExpirationMode.NIST_Unlimited
        DefaultExpirationSettings.Days = 42
        DefaultLockdownSettings.Enabled = True
        DefaultLockdownSettings.DefaultPolicy = True
        DefaultLockdownSettings.TimedLockdownSettings.FailedAttempts = 10
        DefaultLockdownSettings.TimedLockdownSettings.Timeframe = 10
        DefaultLockdownSettings.TimedLockdownSettings.AutoUnlockTime = 10
        DefaultVMSettings.Provider = VMProvider.VirtIO_Guest_Tools
        DefaultNetworkConfiguration.SSID = ""
        DefaultNetworkConfiguration.ConnectWithoutBroadcast = False
        DefaultNetworkConfiguration.Authentication = WiFiAuthenticationMode.WPA2_PSK
        DefaultNetworkConfiguration.Password = ""


        SelectedLanguage = DefaultLanguage
        SelectedLocale = DefaultLocale
        SelectedKeybIdentifier = DefaultKeybIdentifier
        SelectedGeoId = DefaultGeoId
        SelectedOffset = DefaultOffset
        SelectedDiskConfiguration = DefaultDiskConfiguration
        SelectedKey = GenericKeys(5)
        SelectedExpirationSettings = DefaultExpirationSettings
        SelectedLockdownSettings = DefaultLockdownSettings
        SelectedVMSettings = DefaultVMSettings
        SelectedNetworkConfiguration = DefaultNetworkConfiguration

    End Sub

    Sub DetectDotNetRuntime(SDKVersion As String, RuntimeVersion As String)
        If Not Directory.Exists(Path.Combine(Application.StartupPath, "Tools\UnattendGen")) Then
            DotNetRuntimeSupported = False
            Exit Sub
        End If
        If Directory.Exists(Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained")) Then
            ' Self-contained version detected
            DotNetRuntimeSupported = True
            PreferSelfContained = True
            Exit Sub
        End If
        If Not Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet")) Then
            DotNetRuntimeSupported = False
            Exit Sub
        End If
        If Directory.Exists(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet\sdk", SDKVersion)) Then
            ' .NET SDK exists, skip further checks
            DotNetRuntimeSupported = True
            Exit Sub
        End If
        If My.Computer.FileSystem.GetDirectories(Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "dotnet\shared\Microsoft.NETCore.App"), FileIO.SearchOption.SearchTopLevelOnly, RuntimeVersion & "*").Count > 0 Then
            ' .NET Runtime exists, skip further checks
            DotNetRuntimeSupported = True
            Exit Sub
        End If
    End Sub

    Private Sub NewUnattendWiz_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        If MainForm.BackColor = Color.FromArgb(48, 48, 48) Then
            BackColor = Color.FromArgb(31, 31, 31)
            ForeColor = Color.White
            StepsTreeView.BackColor = Color.FromArgb(31, 31, 31)
        ElseIf MainForm.BackColor = Color.FromArgb(239, 239, 242) Then
            BackColor = Color.FromArgb(238, 238, 242)
            ForeColor = Color.Black
            StepsTreeView.BackColor = Color.FromArgb(238, 238, 242)
        End If
        ComboBox1.BackColor = BackColor
        ComboBox2.BackColor = BackColor
        ComboBox3.BackColor = BackColor
        ComboBox4.BackColor = BackColor
        ComboBox5.BackColor = BackColor
        ComboBox6.BackColor = BackColor
        ComboBox7.BackColor = BackColor
        ComboBox8.BackColor = BackColor
        ComboBox9.BackColor = BackColor
        ComboBox10.BackColor = BackColor
        ComboBox11.BackColor = BackColor
        ComboBox12.BackColor = BackColor
        ComboBox13.BackColor = BackColor
        ListBox1.BackColor = BackColor
        ListBox2.BackColor = BackColor
        TextBox1.BackColor = BackColor
        TextBox2.BackColor = BackColor
        TextBox3.BackColor = BackColor
        TextBox4.BackColor = BackColor
        TextBox5.BackColor = BackColor
        TextBox6.BackColor = BackColor
        TextBox7.BackColor = BackColor
        TextBox8.BackColor = BackColor
        TextBox9.BackColor = BackColor
        TextBox10.BackColor = BackColor
        TextBox11.BackColor = BackColor
        TextBox12.BackColor = BackColor
        TextBox13.BackColor = BackColor
        TextBox14.BackColor = BackColor
        TextBox15.BackColor = BackColor
        TextBox17.BackColor = BackColor
        TextBox18.BackColor = BackColor
        NumericUpDown1.BackColor = BackColor
        NumericUpDown2.BackColor = BackColor
        NumericUpDown3.BackColor = BackColor
        NumericUpDown4.BackColor = BackColor
        NumericUpDown5.BackColor = BackColor
        NumericUpDown6.BackColor = BackColor
        NumericUpDown7.BackColor = BackColor
        NumericUpDown8.BackColor = BackColor
        GroupBox1.BackColor = BackColor
        ComboBox1.ForeColor = ForeColor
        ComboBox2.ForeColor = ForeColor
        ComboBox3.ForeColor = ForeColor
        ComboBox4.ForeColor = ForeColor
        ComboBox5.ForeColor = ForeColor
        ComboBox6.ForeColor = ForeColor
        ComboBox7.ForeColor = ForeColor
        ComboBox8.ForeColor = ForeColor
        ComboBox9.ForeColor = ForeColor
        ComboBox10.ForeColor = ForeColor
        ComboBox11.ForeColor = ForeColor
        ComboBox12.ForeColor = ForeColor
        ComboBox13.ForeColor = ForeColor
        ListBox1.ForeColor = ForeColor
        ListBox2.ForeColor = ForeColor
        TextBox1.ForeColor = ForeColor
        TextBox2.ForeColor = ForeColor
        TextBox3.ForeColor = ForeColor
        TextBox4.ForeColor = ForeColor
        TextBox5.ForeColor = ForeColor
        TextBox6.ForeColor = ForeColor
        TextBox7.ForeColor = ForeColor
        TextBox8.ForeColor = ForeColor
        TextBox9.ForeColor = ForeColor
        TextBox10.ForeColor = ForeColor
        TextBox11.ForeColor = ForeColor
        TextBox12.ForeColor = ForeColor
        TextBox13.ForeColor = ForeColor
        TextBox14.ForeColor = ForeColor
        TextBox15.ForeColor = ForeColor
        TextBox17.ForeColor = ForeColor
        TextBox18.ForeColor = ForeColor
        NumericUpDown1.ForeColor = ForeColor
        NumericUpDown2.ForeColor = ForeColor
        NumericUpDown3.ForeColor = ForeColor
        NumericUpDown4.ForeColor = ForeColor
        NumericUpDown5.ForeColor = ForeColor
        NumericUpDown6.ForeColor = ForeColor
        NumericUpDown7.ForeColor = ForeColor
        NumericUpDown8.ForeColor = ForeColor
        GroupBox1.ForeColor = ForeColor
        Dim handle As IntPtr = MainForm.GetWindowHandle(Me)
        If MainForm.IsWindowsVersionOrGreater(10, 0, 18362) Then MainForm.EnableDarkTitleBar(handle, MainForm.BackColor = Color.FromArgb(48, 48, 48))

        SidePanel.BackColor = BackColor
        StepsTreeView.ForeColor = ForeColor
        PictureBox2.Image = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), My.Resources.editor_mode_select, My.Resources.editor_mode)
        ' Fill in font combinations
        FontFamilyTSCB.Items.Clear()
        For Each fntFamily As FontFamily In FontFamily.Families
            FontFamilyTSCB.Items.Add(fntFamily.Name)
        Next
        InitScintilla("Consolas", 11)
        StepsTreeView.ExpandAll()

        FontFamilyTSCB.SelectedItem = "Consolas"
        SetNodeColors(StepsTreeView.Nodes, BackColor, ForeColor)

        DefaultContents = Scintilla1.Text

        SetDefaultSettings()
        ' System language
        If File.Exists(Application.StartupPath & "\AutoUnattend\ImageLanguage.xml") Then
            ImageLanguages = ImageLanguage.LoadItems(Application.StartupPath & "\AutoUnattend\ImageLanguage.xml")
            If ImageLanguages IsNot Nothing Then
                For Each imgLang As ImageLanguage In ImageLanguages
                    ComboBox1.Items.Add(imgLang.DisplayName)
                Next
                If ComboBox1.SelectedItem = Nothing Then ComboBox1.SelectedItem = DefaultLanguage.DisplayName
            End If
        End If
        ' System locale
        If File.Exists(Application.StartupPath & "\AutoUnattend\UserLocale.xml") Then
            UserLocales = UserLocale.LoadItems(Application.StartupPath & "\AutoUnattend\UserLocale.xml")
            If UserLocales IsNot Nothing Then
                For Each userLoc As UserLocale In UserLocales
                    ComboBox2.Items.Add(userLoc.DisplayName)
                Next
                If ComboBox2.SelectedItem = Nothing Then ComboBox2.SelectedItem = DefaultLocale.DisplayName
            End If
        End If
        ' Keyboard layout/IME
        If File.Exists(Application.StartupPath & "\AutoUnattend\KeyboardIdentifier.xml") Then
            KeyboardIdentifiers = KeyboardIdentifier.LoadItems(Application.StartupPath & "\AutoUnattend\KeyboardIdentifier.xml")
            If KeyboardIdentifiers IsNot Nothing Then
                For Each keyb As KeyboardIdentifier In KeyboardIdentifiers
                    ComboBox3.Items.Add(keyb.DisplayName)
                Next
                If ComboBox3.SelectedItem = Nothing Then ComboBox3.SelectedItem = DefaultKeybIdentifier.DisplayName
            End If
        End If
        ' Home location
        If File.Exists(Application.StartupPath & "\AutoUnattend\GeoId.xml") Then
            GeoIds = GeoId.LoadItems(Application.StartupPath & "\AutoUnattend\GeoId.xml")
            If GeoIds IsNot Nothing Then
                For Each Geo As GeoId In GeoIds
                    ComboBox4.Items.Add(Geo.DisplayName)
                Next
                If ComboBox4.SelectedItem = Nothing Then ComboBox4.SelectedItem = DefaultGeoId.DisplayName
            End If
        End If
        ' Time offsets
        If File.Exists(Application.StartupPath & "\AutoUnattend\TimeOffset.xml") Then
            TimeOffsets = TimeOffset.LoadItems(Application.StartupPath & "\AutoUnattend\TimeOffset.xml")
            If TimeOffsets IsNot Nothing Then
                For Each Offset As TimeOffset In TimeOffsets
                    ComboBox5.Items.Add(Offset.DisplayName)
                Next
                If ComboBox5.SelectedItem = Nothing Then ComboBox5.SelectedItem = DefaultOffset.DisplayName
            End If
        End If
        ' System components
        If File.Exists(Application.StartupPath & "\AutoUnattend\Component.xml") Then
            SystemComponents = Component.LoadItems(Application.StartupPath & "\AutoUnattend\Component.xml")
            DefaultSystemComponents = Component.LoadItems(Application.StartupPath & "\AutoUnattend\Component.xml")
            If SystemComponents IsNot Nothing Then
                For Each SystemComponent As Component In SystemComponents
                    ListBox2.Items.Add(SystemComponent.Id)
                Next
            End If
        End If
        ListBox1.SelectedIndex = 1
        ChangePage(UnattendedWizardPage.Page.WelcomePage)
        VerifyInPages.AddRange(New UnattendedWizardPage.Page() {UnattendedWizardPage.Page.SysConfigPage, UnattendedWizardPage.Page.DiskConfigPage, UnattendedWizardPage.Page.ProductKeyPage, UnattendedWizardPage.Page.UserAccountsPage, UnattendedWizardPage.Page.NetworkConnectionsPage})
        TimeZonePageTimer.Enabled = True
        ' Modify script contents of disk config for sample DP Script
        SelectedDiskConfiguration.DiskPartScriptConfig.ScriptContents = Scintilla2.Text
        ' Set PRO edition
        If ComboBox6.SelectedItem = Nothing Then ComboBox6.SelectedItem = "Pro"
        ' Set default auth tech to WPA2
        If ComboBox13.SelectedItem = Nothing Then ComboBox13.SelectedItem = "WPA2-PSK"

        ' Detect .NET runtimes/SDKs
        DetectDotNetRuntime("9.0.100", "9.0")
        If Not DotNetRuntimeSupported Then
            If MsgBox("This wizard requires the .NET 9 Runtime to be installed to use the built-in version of the generator program. You can download it from:" & CrLf & CrLf & "dotnet.microsoft.com" & CrLf & CrLf & "If you don't want to download .NET, you can download the self-contained version of the generator program. Downloading it will take some time, depending on your network connection speed." & CrLf & CrLf & "Do you want to use the self-contained version?", vbYesNo + vbQuestion, ".NET Runtime missing") = Windows.Forms.DialogResult.Yes Then
                ExpressPanelFooter.Enabled = False
                UnattendGenBW.RunWorkerAsync()
            Else
                Close()
            End If
        Else
            UGNotify.Visible = False
        End If

        ' Detect presence of Windows SIM
        If File.Exists(Path.Combine(Environment.GetFolderPath(If(Environment.Is64BitOperatingSystem, Environment.SpecialFolder.ProgramFilesX86, Environment.SpecialFolder.ProgramFiles)),
                                    "Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\WSIM\x86\imgmgr.exe")) Then
            LinkLabel6.Enabled = True
        Else
            LinkLabel6.Enabled = False
        End If
    End Sub

    Sub ReloadSettings()
        ' Restore regional configuration
        ComboBox1.SelectedItem = DefaultLanguage.DisplayName
        ComboBox2.SelectedItem = DefaultLocale.DisplayName
        ComboBox3.SelectedItem = DefaultKeybIdentifier.DisplayName
        ComboBox4.SelectedItem = DefaultGeoId.DisplayName
        ' Restore basic system configuration
        ListBox1.SelectedIndex = 1
        Win11Config.LabConfig_BypassRequirements = False
        Win11Config.OOBE_BypassNRO = False
        CheckBox1.Checked = False
        CheckBox2.Checked = False
        CheckBox3.Checked = True
        TextBox1.Text = ""
        ' Restore time zone
        ComboBox5.SelectedItem = DefaultOffset.DisplayName
        RadioButton1.Checked = True
        ' Restore disk configuration
        CheckBox4.Checked = True
        RadioButton5.Checked = True
        RadioButton7.Checked = True
        NumericUpDown1.Value = 300
        CheckBox5.Checked = True
        RadioButton9.Checked = True
        NumericUpDown2.Value = 1000
        Scintilla2.Text = My.Resources.DefaultDiskPartConfig
        RadioButton11.Checked = True
        NumericUpDown3.Value = 0
        NumericUpDown4.Value = 3
        SelectedDiskConfiguration = DefaultDiskConfiguration
        ' Restore product key
        RadioButton13.Checked = True
        ComboBox6.SelectedItem = "Pro"
        TextBox3.Text = ""
        ' Restore user accounts
        CheckBox6.Checked = True
        TextBox4.Text = "Admin"
        TextBox6.Text = ""
        TextBox8.Text = ""
        TextBox9.Text = ""
        TextBox11.Text = ""
        TextBox12.Text = ""
        TextBox14.Text = ""
        TextBox15.Text = ""
        TextBox17.Text = ""
        TextBox18.Text = ""
        CheckBox8.Checked = False
        CheckBox9.Checked = False
        CheckBox10.Checked = False
        CheckBox11.Checked = False
        ComboBox7.SelectedIndex = 0
        ComboBox9.SelectedIndex = 1
        ComboBox10.SelectedIndex = 1
        ComboBox11.SelectedIndex = 1
        ComboBox12.SelectedIndex = 1
        CheckBox12.Checked = False
        RadioButton15.Checked = True
        TextBox5.Text = ""
        CheckBox7.Checked = True
        CheckBox18.Checked = False
        ' Restore password expiration
        RadioButton17.Checked = True
        RadioButton19.Checked = True
        NumericUpDown5.Value = 10
        ' Restore Account lockout
        CheckBox13.Checked = False
        RadioButton21.Checked = True
        NumericUpDown6.Value = 10
        NumericUpDown7.Value = 10
        NumericUpDown8.Value = 10
        ' Restore VM support
        ComboBox8.SelectedIndex = 2
        RadioButton24.Checked = True
        ' Restore network settings
        CheckBox14.Checked = True
        RadioButton25.Checked = True
        TextBox7.Text = ""
        CheckBox15.Checked = False
        ComboBox13.SelectedIndex = 1
        TextBox10.Text = ""
        ' Restore system telemetry
        CheckBox16.Checked = False
        RadioButton26.Checked = True
        ' Restore default selections for components
        SystemComponents = DefaultSystemComponents

        ' Restore variables
        UserAccountsList.Clear()
        SetDefaultSettings()
    End Sub

    Sub SelectTreeNode(NodeIndex As Integer)
        StepsTreeView.SelectedNode = StepsTreeView.Nodes(NodeIndex)
        StepsTreeView.Refresh()
    End Sub

    Sub ChangePage(NewPage As UnattendedWizardPage.Page)
        If NewPage > CurrentWizardPage.WizardPage AndAlso VerifyInPages.Contains(CurrentWizardPage.WizardPage) Then
            If Not VerifyOptionsInPage(CurrentWizardPage.WizardPage) Then Exit Sub
        ElseIf NewPage > CurrentWizardPage.WizardPage AndAlso NewPage = UnattendedWizardPage.Page.ReviewPage Then
            ShowSettingOverview()
        End If
        WelcomePanel.Visible = (NewPage = UnattendedWizardPage.Page.WelcomePage)
        RegionalSettingsPanel.Visible = (NewPage = UnattendedWizardPage.Page.RegionalPage)
        SysConfigPanel.Visible = (NewPage = UnattendedWizardPage.Page.SysConfigPage)
        TimeZonePanel.Visible = (NewPage = UnattendedWizardPage.Page.TimeZonePage)
        DiskConfigurationPanel.Visible = (NewPage = UnattendedWizardPage.Page.DiskConfigPage)
        ProductKeyPanel.Visible = (NewPage = UnattendedWizardPage.Page.ProductKeyPage)
        UserAccountPanel.Visible = (NewPage = UnattendedWizardPage.Page.UserAccountsPage)
        PWExpirationPanel.Visible = (NewPage = UnattendedWizardPage.Page.PWExpirationPage)
        AccountLockoutPanel.Visible = (NewPage = UnattendedWizardPage.Page.AccountLockoutPage)
        VirtualMachinePanel.Visible = (NewPage = UnattendedWizardPage.Page.VirtualMachinePage)
        NetworkConnectionPanel.Visible = (NewPage = UnattendedWizardPage.Page.NetworkConnectionsPage)
        SystemTelemetryPanel.Visible = (NewPage = UnattendedWizardPage.Page.SystemTelemetryPage)
        PostInstallPanel.Visible = (NewPage = UnattendedWizardPage.Page.PostInstallPage)
        ComponentPanel.Visible = (NewPage = UnattendedWizardPage.Page.ComponentPage)
        FinalReviewPanel.Visible = (NewPage = UnattendedWizardPage.Page.ReviewPage)
        UnattendProgressPanel.Visible = (NewPage = UnattendedWizardPage.Page.ProgressPage)
        FinishPanel.Visible = (NewPage = UnattendedWizardPage.Page.FinishPage)
        CurrentWizardPage.WizardPage = NewPage
        Next_Button.Enabled = (Not NewPage <> UnattendedWizardPage.Page.FinishPage) OrElse (Not NewPage + 1 >= UnattendedWizardPage.PageCount)
        Cancel_Button.Enabled = Not (NewPage = UnattendedWizardPage.Page.FinishPage)
        Back_Button.Enabled = Not (NewPage = UnattendedWizardPage.Page.WelcomePage) And Not (NewPage = UnattendedWizardPage.Page.FinishPage)

        Next_Button.Text = If(NewPage = UnattendedWizardPage.Page.FinishPage, "Close", "Next")

        ' Select tree nodes according to page
        Select Case CurrentWizardPage.WizardPage
            Case UnattendedWizardPage.Page.WelcomePage
                SelectTreeNode(0)
            Case UnattendedWizardPage.Page.RegionalPage
                SelectTreeNode(1)
            Case UnattendedWizardPage.Page.SysConfigPage
                SelectTreeNode(2)
            Case UnattendedWizardPage.Page.TimeZonePage
                SelectTreeNode(3)
            Case UnattendedWizardPage.Page.DiskConfigPage
                SelectTreeNode(4)
            Case UnattendedWizardPage.Page.ProductKeyPage
                SelectTreeNode(5)
            Case UnattendedWizardPage.Page.UserAccountsPage, UnattendedWizardPage.Page.PWExpirationPage, UnattendedWizardPage.Page.AccountLockoutPage
                SelectTreeNode(6)
            Case UnattendedWizardPage.Page.VirtualMachinePage
                SelectTreeNode(7)
            Case UnattendedWizardPage.Page.NetworkConnectionsPage
                SelectTreeNode(8)
            Case UnattendedWizardPage.Page.SystemTelemetryPage
                SelectTreeNode(9)
            Case UnattendedWizardPage.Page.PostInstallPage
                SelectTreeNode(10)
            Case UnattendedWizardPage.Page.ComponentPage
                SelectTreeNode(11)
            Case UnattendedWizardPage.Page.ReviewPage, UnattendedWizardPage.Page.ProgressPage, UnattendedWizardPage.Page.FinishPage
                SelectTreeNode(12)
        End Select

        ' Change sizes of controls if the normal resize event does not work
        AutoDiskConfigPanel.Width = ManualPartPanel.Width - (AutoDiskConfigPanel.Margin.Left * 2) - 4
        DiskPartPanel.Width = ManualPartPanel.Width - (DiskPartPanel.Margin.Left * 2) - 4
        GroupBox1.Width = ManualAccountPanel.Width - (GroupBox1.Margin.Left * 2) - 4
        AccountsPanel.Width = UserAccountListing.Width
        UserAccountListing.Width = ManualAccountPanel.Width - (UserAccountListing.Margin.Left * 2) - 4
        WirelessNetworkSettingsPanel.Width = ManualNetworkConfigPanel.Width - (WirelessNetworkSettingsPanel.Margin.Left * 2) - 4

        ExpressPanelFooter.Enabled = Not (CurrentWizardPage.WizardPage = UnattendedWizardPage.Page.ProgressPage)
        If CurrentWizardPage.WizardPage = UnattendedWizardPage.Page.ProgressPage Then
            ' Detect if a project has been loaded
            If MainForm.isProjectLoaded And Not (MainForm.OnlineManagement Or MainForm.OfflineManagement) Then
                SaveFileDialog1.InitialDirectory = Path.Combine(MainForm.projPath, "unattend_xml")
            Else
                SaveFileDialog1.InitialDirectory = ""
            End If
            SaveFileDialog1.FileName = "autounattend_" & Now.ToString().Replace("/", "-").Trim().Replace(":", "-").Trim() & ".xml"
            SaveFileDialog1.ShowDialog()
            UnattendGeneratorBW.RunWorkerAsync()
        ElseIf CurrentWizardPage.WizardPage = UnattendedWizardPage.Page.ReviewPage Then
            SaveTarget = ""
        End If
    End Sub

    Function VerifyOptionsInPage(WizardPage As UnattendedWizardPage.Page) As Boolean
        Select Case WizardPage
            Case UnattendedWizardPage.Page.SysConfigPage
                If ListBox1.SelectedItems.Count = 0 Then
                    MessageBox.Show("Please select an architecture and try again", "Validation error")
                    Return False
                End If
                If Not PCName.DefaultName Then
                    Dim testerPC As ComputerName = ComputerNameValidator.ValidateComputerName(TextBox1.Text)
                    If Not testerPC.Valid AndAlso testerPC.ErrorMessage <> "" Then
                        MessageBox.Show(testerPC.ErrorMessage, "Computer name error")
                        Return False
                    End If
                End If
            Case UnattendedWizardPage.Page.DiskConfigPage
                If Not DiskConfigurationInteractive AndAlso SelectedDiskConfiguration.DiskConfigMode = DiskConfigurationMode.DiskPart AndAlso Scintilla2.Text = "" Then
                    MessageBox.Show("Please enter the contents of the DiskPart script and try again. You can also use a script file", "DiskPart Script error")
                    Return False
                End If
            Case UnattendedWizardPage.Page.ProductKeyPage
                If Not GenericChosen Then
                    If TextBox3.Text = "" Then
                        MessageBox.Show("Please type a product key and try again", "Product Key error")
                        Return False
                    ElseIf TextBox3.Text <> "" And TextBox3.Text.Length <> 29 Then
                        MessageBox.Show("Please type all of the product key and try again", "Product Key error")
                        Return False
                    ElseIf TextBox3.Text <> "" And TextBox3.Text.Length = 29 Then
                        Dim pKey As ProductKey = ProductKeyValidator.ValidateProductKey(TextBox3.Text)
                        If Not pKey.Valid Then
                            MessageBox.Show("The product key entered:" & CrLf & CrLf & TextBox3.Text & CrLf & CrLf & "is ill-formed. Please type it again", "Product Key error")
                            Return False
                        End If
                    End If
                End If
            Case UnattendedWizardPage.Page.UserAccountsPage
                Dim validationResults As UserValidationResults = UserValidator.ValidateUsers(UserAccountsList, PCName)
                If Not UserAccountsInteractive AndAlso Not MicrosoftAccountInteractive AndAlso Not validationResults.IsValid Then
                    MessageBox.Show("There is a problem with one or more of the users specified. Here are the reasons why:" & CrLf & CrLf & validationResults.ValidationErrorReason & CrLf & CrLf & "Try again after fixing the aforementioned problems", "User Accounts error", MessageBoxButtons.OK, MessageBoxIcon.Exclamation)
                    Return False
                End If
                Dim invalidChars As Char() = {"/", "\", "[", "]", ":", ";", "|", "=", ",", "+", "*", "?", "<", ">"}
                If Not UserAccountsInteractive AndAlso Not MicrosoftAccountInteractive Then
                    Dim AtLeastOneAdmin As Boolean = False
                    If UserAccountsList.Count > 0 Then
                        For Each UserAccount As User In UserAccountsList
                            UserAccount.Name = New String(UserAccount.Name.Where(Function(c) Not invalidChars.Contains(c)).ToArray())
                            If UserAccount.Group = UserGroup.Administrators Then
                                AtLeastOneAdmin = True
                                Exit For
                            End If
                        Next
                    End If
                    If Not AtLeastOneAdmin Then
                        MessageBox.Show("At least one account must be part of the Administrators user group. Please configure the user groups accordingly and try again", "User Accounts error")
                        Return False
                    End If
                End If
            Case UnattendedWizardPage.Page.NetworkConnectionsPage
                If Not NetworkConfigInteractive AndAlso Not NetworkConfigManualSkip AndAlso Not WirelessValidator.ValidateWiFi(SelectedNetworkConfiguration) Then
                    MessageBox.Show("There is a problem with the specified wireless settings. Make sure that you have specified a network name and try again", "Wireless Networks error")
                    Return False
                End If
        End Select
        Return True
    End Function

    Sub ShowSettingOverview()
        TextBox13.Clear()
        ' Display settings in the following order:
        TextBox13.Text = "Current configurations for the unattended answer file:" & CrLf
        ' 1. -- REGIONAL CONFIGURATION
        TextBox13.AppendText("Regional settings: " & If(RegionalInteractive, "configured during setup" & CrLf, CrLf))
        If Not RegionalInteractive Then
            TextBox13.AppendText("- System language: " & SelectedLanguage.DisplayName & CrLf &
                                 "- System locale: " & SelectedLocale.DisplayName & CrLf &
                                 "- Keyboard/IME: " & SelectedKeybIdentifier.DisplayName & CrLf &
                                 "- Home location: " & SelectedGeoId.DisplayName & CrLf)
        End If
        ' 2. -- BASIC SYSTEM CONFIGURATION
        TextBox13.AppendText("Basic system configuration: " & CrLf &
                             "- Processor architecture: " & Utilities.Casters.CastDismArchitecture(SelectedArchitecture) & CrLf &
                             "- Windows 11 Settings:" & CrLf &
                             "    - Bypass System Requirements? " & If(Win11Config.LabConfig_BypassRequirements, "Yes", "No") & CrLf &
                             "    - Bypass Mandatory Network Connection? " & If(Win11Config.OOBE_BypassNRO, "Yes", "No") & CrLf &
                             "- Computer name: " & If(PCName.DefaultName, "random by Windows", PCName.Name) & CrLf)
        ' 3. -- TIME ZONE
        TextBox13.AppendText("Time zone configuration: " & If(TimeOffsetInteractive, "based on regional settings" & CrLf, CrLf))
        If Not TimeOffsetInteractive Then
            TextBox13.AppendText("- Time zone: " & SelectedOffset.DisplayName & CrLf)
        End If
        ' 4. -- DISK CONFIGURATION
        TextBox13.AppendText("Disk configuration: " & If(DiskConfigurationInteractive, "configured during setup" & CrLf, CrLf))
        If Not DiskConfigurationInteractive Then
            TextBox13.AppendText("- Disk configuration mode: " & If(SelectedDiskConfiguration.DiskConfigMode = DiskConfigurationMode.AutoDisk0, "automatically configure Disk 0", "configure disks with a DiskPart script") & CrLf)
            Select Case SelectedDiskConfiguration.DiskConfigMode
                Case DiskConfigurationMode.AutoDisk0
                    TextBox13.AppendText("    - Partition table: " & If(SelectedDiskConfiguration.PartStyle = PartitionStyle.GPT, "GPT (UEFI)", "MBR (BIOS/CSM)") & CrLf)
                    If SelectedDiskConfiguration.PartStyle = PartitionStyle.GPT Then
                        TextBox13.AppendText("      - EFI System Partition Size: " & SelectedDiskConfiguration.ESPSize & " MB" & CrLf)
                    End If
                    TextBox13.AppendText("    - Install a Recovery Environment? " & If(SelectedDiskConfiguration.InstallRecEnv, "Yes", "No") & CrLf)
                    If SelectedDiskConfiguration.InstallRecEnv Then
                        TextBox13.AppendText("      - Location of the Recovery Environment: " & If(SelectedDiskConfiguration.RecEnvPartition = RecoveryEnvironmentLocation.WinREPartition, "Recovery partition", "Windows partition") & CrLf)
                        If SelectedDiskConfiguration.RecEnvPartition = RecoveryEnvironmentLocation.WinREPartition Then
                            TextBox13.AppendText("        - Recovery Partition Size: " & SelectedDiskConfiguration.RecEnvSize & " MB" & CrLf)
                        End If
                    End If
                Case DiskConfigurationMode.DiskPart
                    TextBox13.AppendText("    - Action to be performed after disk configuration: " & If(SelectedDiskConfiguration.DiskPartScriptConfig.AutomaticInstall, "install to first available partition with enough space and no installations", "install to specific disk") & CrLf)
                    If Not SelectedDiskConfiguration.DiskPartScriptConfig.AutomaticInstall Then
                        TextBox13.AppendText("      - Target disk/partition: disk " & SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.DiskNum & ", partition " & SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.PartNum & CrLf)
                    End If
            End Select
        End If
        ' 5. -- PRODUCT KEY
        TextBox13.AppendText("Product key: " & If(GenericChosen, "generic" & CrLf, "custom" & CrLf) &
                             "- Key: " & SelectedKey.Key & CrLf)
        ' 6. -- USER ACCOUNTS
        TextBox13.AppendText("User account settings: " & If(UserAccountsInteractive, "configured during setup" & CrLf, CrLf))
        If Not UserAccountsInteractive And Not MicrosoftAccountInteractive Then
            For Each UserAccount As User In UserAccountsList
                TextBox13.AppendText("- Account " & UserAccountsList.IndexOf(UserAccount) + 1 & "? " & If(UserAccount.Enabled, "Yes", "No") & CrLf)
                If UserAccount.Enabled Then
                    TextBox13.AppendText("    - Name: " & UserAccount.Name & CrLf &
                                         "    - Password: " & UserAccount.Password & CrLf &
                                         "    - Group: " & If(UserAccount.Group = UserGroup.Administrators, "Administrators", "Users") & CrLf)
                End If
            Next
            ' First logon settings
            TextBox13.AppendText("- Log on as an Administrator account? " & If(AutoLogon.EnableAutoLogon, "Yes", "No") & CrLf)
            If AutoLogon.EnableAutoLogon Then
                TextBox13.AppendText("    - Administrator account: " & If(AutoLogon.LogonMode = AutoLogonMode.FirstAdmin, "first admin account created", "built-in Administrator account") & CrLf)
                If AutoLogon.LogonMode = AutoLogonMode.WindowsAdmin Then
                    TextBox13.AppendText("      - Password for built-in administrator: " & AutoLogon.LogonPassword & CrLf)
                End If
            End If
            TextBox13.AppendText("- Obscure passwords with Base64? " & If(PasswordObfuscate, "Yes", "No") & CrLf)
        ElseIf (Not UserAccountsInteractive) And MicrosoftAccountInteractive Then
            TextBox13.AppendText("- The target system will ask for a Microsoft account" & CrLf)
        End If
        TextBox13.AppendText("Password expiration policy: " & If(SelectedExpirationSettings.Mode = PasswordExpirationMode.NIST_Limited, "enabled" & CrLf, "disabled" & CrLf))
        If SelectedExpirationSettings.Mode = PasswordExpirationMode.NIST_Limited Then
            TextBox13.AppendText("- Expiration policy mode: " & If(SelectedExpirationSettings.WindowsDefault, "Windows default (42 days)", "custom") & CrLf)
            If Not SelectedExpirationSettings.WindowsDefault Then
                TextBox13.AppendText("    - Expiration period: " & SelectedExpirationSettings.Days & " days" & CrLf)
            End If
        End If
        TextBox13.AppendText("Account Lockout policy status: " & If(SelectedLockdownSettings.Enabled, "enabled" & CrLf, "disabled" & CrLf))
        If SelectedLockdownSettings.Enabled Then
            TextBox13.AppendText("- Account Lockout policies: " & If(SelectedLockdownSettings.DefaultPolicy, "default", "custom") & CrLf)
            If Not SelectedLockdownSettings.DefaultPolicy Then
                TextBox13.AppendText("    - After " & SelectedLockdownSettings.TimedLockdownSettings.FailedAttempts & " failed attempts within " & SelectedLockdownSettings.TimedLockdownSettings.Timeframe & " minutes, unlock account after " & SelectedLockdownSettings.TimedLockdownSettings.AutoUnlockTime & " minutes" & CrLf)
            End If
        End If
        ' 7. -- VIRTUAL MACHINE SUPPORT
        TextBox13.AppendText("Virtual Machine Support: " & If(VirtualMachineSupported, "enabled" & CrLf, "disabled" & CrLf))
        If VirtualMachineSupported Then
            Select Case SelectedVMSettings.Provider
                Case VMProvider.VirtualBox_GAs
                    TextBox13.AppendText("- Selected Hypervisor: Oracle VM VirtualBox (VirtualBox Guest Additions)" & CrLf)
                Case VMProvider.VMware_Tools
                    TextBox13.AppendText("- Selected Hypervisor: VMware (VMware Tools)" & CrLf)
                Case VMProvider.VirtIO_Guest_Tools
                    TextBox13.AppendText("- Selected Hypervisor: QEMU/Proxmox VE/etc. (VirtIO Guest Tools)" & CrLf)
            End Select
        End If
        ' 8. -- WIRELESS NETWORKING
        TextBox13.AppendText("Wireless networking settings: " & If(NetworkConfigInteractive, "configured during setup" & CrLf, CrLf))
        If Not NetworkConfigInteractive Then
            TextBox13.AppendText("- Skip configuration? " & If(NetworkConfigManualSkip, "Yes", "No") & CrLf)
            If Not NetworkConfigManualSkip Then
                TextBox13.AppendText("    - SSID: " & SelectedNetworkConfiguration.SSID & CrLf &
                                     "    - Connect even if not broadcasting? " & If(SelectedNetworkConfiguration.ConnectWithoutBroadcast, "Yes", "No") & CrLf)
                Select Case SelectedNetworkConfiguration.Authentication
                    Case WiFiAuthenticationMode.Open
                        TextBox13.AppendText("    - Authentication mode: open" & CrLf)
                    Case WiFiAuthenticationMode.WPA2_PSK
                        TextBox13.AppendText("    - Authentication mode: WPA2-Personal" & CrLf)
                    Case WiFiAuthenticationMode.WPA3_SAE
                        TextBox13.AppendText("    - Authentication mode: WPA3 (Simultaneous Authentication of Equals)" & CrLf)
                End Select
                TextBox13.AppendText("    - Password: " & New String("*", SelectedNetworkConfiguration.Password.Length) & " (hidden for your security)" & CrLf)
            End If
        End If
        ' 9. -- SYSTEM TELEMETRY
        TextBox13.AppendText("System telemetry settings: " & If(SystemTelemetryInteractive, "configured during setup" & CrLf, CrLf))
        If Not SystemTelemetryInteractive Then
            TextBox13.AppendText("- (Attempt to) disable telemetry? " & If(Not SelectedTelemetrySettings.Enabled, "Yes", "No") & CrLf)
        End If
        ' Post Install Scripts and Component Manager will be added in a future release
        ' 11. -- COMPONENTS
        TextBox13.AppendText("Additional components: " & If(Not AreComponentListsEqual(SystemComponents, DefaultSystemComponents), "", "none") & CrLf)
        If Not AreComponentListsEqual(SystemComponents, DefaultSystemComponents) Then
            FinalComponents = GetComponentDifferences(SystemComponents, DefaultSystemComponents)
            If FinalComponents.Count > 0 Then
                For Each systemComponent As Component In FinalComponents
                    TextBox13.AppendText("- Component name: " & Quote & systemComponent.Id & Quote & CrLf &
                                         "  - Passes:" & CrLf)
                    If systemComponent.Passes.Count > 0 Then
                        For Each systemPass As Pass In systemComponent.Passes
                            TextBox13.AppendText("    - " & Quote & systemPass.Name & Quote & CrLf)
                        Next
                    End If
                Next
            End If
        End If
    End Sub

    Private Sub ExpressPanelTrigger_MouseEnter(sender As Object, e As EventArgs) Handles ExpressPanelTrigger.MouseEnter
        If ExpressPanelContainer.Visible Then
            ExpressPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.HotTrack)
        Else
            ExpressPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(48, 48, 48), Color.Gainsboro)
        End If
    End Sub

    Private Sub ExpressPanelTrigger_MouseLeave(sender As Object, e As EventArgs) Handles ExpressPanelTrigger.MouseLeave
        If ExpressPanelContainer.Visible Then
            ExpressPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        Else
            ExpressPanelTrigger.BackColor = SidePanel.BackColor
        End If
    End Sub

    Private Sub ExpressPanelTrigger_MouseDown(sender As Object, e As MouseEventArgs) Handles ExpressPanelTrigger.MouseDown
        If ExpressPanelContainer.Visible Then
            ExpressPanelTrigger.BackColor = Color.SteelBlue
        Else
            ExpressPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(36, 36, 36), Color.Silver)
        End If
    End Sub

    Private Sub ExpressPanelTrigger_MouseUp(sender As Object, e As MouseEventArgs) Handles ExpressPanelTrigger.MouseUp
        If ExpressPanelContainer.Visible Then
            ExpressPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.HotTrack)
        Else
            ExpressPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(48, 48, 48), Color.Gainsboro)
        End If
    End Sub

    Private Sub ExpressPanelTrigger_Click(sender As Object, e As EventArgs) Handles ExpressPanelTrigger.Click
        IsInExpress = True
        StepsTreeView.Enabled = True
        EditorPanelContainer.Visible = False
        ExpressPanelContainer.Visible = True
        ExpressPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        ExpressPanelTrigger.ForeColor = Color.White
        PictureBox1.Image = My.Resources.express_mode_select
        EditorPanelTrigger.BackColor = SidePanel.BackColor
        EditorPanelTrigger.ForeColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.LightGray, Color.Black)
        PictureBox2.Image = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), My.Resources.editor_mode_select, My.Resources.editor_mode)
        PictureBox3.Image = My.Resources.express_mode_fc
        Label3.Text = "Express mode"
        Label4.Text = "If you haven't created unattended answer files before, use this wizard to create one"
        FooterContainer.Visible = True
    End Sub

    Private Sub EditorPanelTrigger_MouseEnter(sender As Object, e As EventArgs) Handles EditorPanelTrigger.MouseEnter
        If EditorPanelContainer.Visible Then
            EditorPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.HotTrack)
        Else
            EditorPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(48, 48, 48), Color.Gainsboro)
        End If
    End Sub

    Private Sub EditorPanelTrigger_MouseLeave(sender As Object, e As EventArgs) Handles EditorPanelTrigger.MouseLeave
        If EditorPanelContainer.Visible Then
            EditorPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        Else
            EditorPanelTrigger.BackColor = SidePanel.BackColor
        End If
    End Sub

    Private Sub EditorPanelTrigger_MouseDown(sender As Object, e As MouseEventArgs) Handles EditorPanelTrigger.MouseDown
        If EditorPanelContainer.Visible Then
            EditorPanelTrigger.BackColor = Color.SteelBlue
        Else
            EditorPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(36, 36, 36), Color.Silver)
        End If
    End Sub

    Private Sub EditorPanelTrigger_MouseUp(sender As Object, e As MouseEventArgs) Handles EditorPanelTrigger.MouseUp
        If EditorPanelContainer.Visible Then
            EditorPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.HotTrack)
        Else
            EditorPanelTrigger.BackColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(48, 48, 48), Color.Gainsboro)
        End If
    End Sub

    Private Sub EditorPanelTrigger_Click(sender As Object, e As EventArgs) Handles EditorPanelTrigger.Click
        IsInExpress = False
        StepsTreeView.Enabled = False
        EditorPanelContainer.Visible = True
        ExpressPanelContainer.Visible = False
        ExpressPanelTrigger.BackColor = SidePanel.BackColor
        ExpressPanelTrigger.ForeColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.LightGray, Color.Black)
        PictureBox1.Image = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), My.Resources.express_mode_select, My.Resources.express_mode)
        EditorPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        EditorPanelTrigger.ForeColor = Color.White
        PictureBox2.Image = My.Resources.editor_mode_select
        PictureBox3.Image = My.Resources.editor_mode_fc
        Label3.Text = "Editor mode"
        Label4.Text = "Create your unattended answer files from scratch and save them anywhere"
        FooterContainer.Visible = False
    End Sub

    Private Sub Back_Button_Click(sender As Object, e As EventArgs) Handles Back_Button.Click
        ChangePage(CurrentWizardPage.WizardPage - 1)
    End Sub

    Private Sub Next_Button_Click(sender As Object, e As EventArgs) Handles Next_Button.Click
        If CurrentWizardPage.WizardPage = UnattendedWizardPage.Page.FinishPage Then
            Close()
        Else
            ChangePage(CurrentWizardPage.WizardPage + 1)
        End If
    End Sub

    Private Sub Cancel_Button_Click(sender As Object, e As EventArgs) Handles Cancel_Button.Click
        Close()
    End Sub

    Private Sub FontChange(sender As Object, e As EventArgs) Handles FontFamilyTSCB.SelectedIndexChanged, FontSizeTSCB.SelectedIndexChanged
        ' Change Scintilla editor font
        InitScintilla(FontFamilyTSCB.SelectedItem, FontSizeTSCB.SelectedItem)
    End Sub

    Private Sub ToolStripButton5_Click(sender As Object, e As EventArgs) Handles ToolStripButton5.Click
        If ToolStripButton5.Checked Then
            ToolStripButton5.Checked = False
        Else
            ToolStripButton5.Checked = True
        End If
        Scintilla1.WrapMode = If(ToolStripButton5.Checked, WrapMode.Word, WrapMode.None)
    End Sub

    Sub SetNodeColors(nodes As TreeNodeCollection, bg As Color, fg As Color)
        For Each node As TreeNode In nodes
            node.BackColor = BackColor
            node.ForeColor = ForeColor
            SetNodeColors(node.Nodes, BackColor, ForeColor)
        Next
    End Sub

    Private Sub RadioButton1_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton1.CheckedChanged
        RegionalInteractive = Not RadioButton1.Checked
        RegionalSettings.Enabled = RadioButton1.Checked
        Label10.Enabled = Not RadioButton1.Checked
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        SelectedLanguage = ImageLanguages(ComboBox1.SelectedIndex)
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        SelectedLocale = UserLocales(ComboBox2.SelectedIndex)
    End Sub

    Private Sub ComboBox3_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox3.SelectedIndexChanged
        SelectedKeybIdentifier = KeyboardIdentifiers(ComboBox3.SelectedIndex)
    End Sub

    Private Sub ComboBox4_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox4.SelectedIndexChanged
        SelectedGeoId = GeoIds(ComboBox4.SelectedIndex)
    End Sub

    Private Sub ListBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox1.SelectedIndexChanged
        Select Case ListBox1.SelectedIndex
            Case 0
                SelectedArchitecture = DismProcessorArchitecture.Intel
            Case 1
                SelectedArchitecture = DismProcessorArchitecture.AMD64
            Case 2
                SelectedArchitecture = DismProcessorArchitecture.ARM64
        End Select
        ' Disable Windows 11 settings for x86
        WinSVSettingsPanel.Enabled = Not (SelectedArchitecture = DismProcessorArchitecture.Intel)
    End Sub

    Private Sub CheckBox3_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox3.CheckedChanged
        PCName.DefaultName = CheckBox3.Checked
        ComputerNamePanel.Enabled = Not CheckBox3.Checked
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        Win11Config.LabConfig_BypassRequirements = CheckBox1.Checked
    End Sub

    Private Sub CheckBox2_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox2.CheckedChanged
        Win11Config.OOBE_BypassNRO = CheckBox2.Checked
    End Sub

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged
        Try
            If New StackFrame(6).GetMethod().Name = "ReloadSettings" Then
                Exit Sub
            End If
        Catch ex As Exception
            ' Continue the method
        End Try
        ' Hold default value for now
        Dim defVal As Boolean = False
        defVal = PCName.DefaultName
        PCName = ComputerNameValidator.ValidateComputerName(TextBox1.Text)
        PCName.DefaultName = defVal
        If Not PCName.Valid AndAlso PCName.ErrorMessage <> "" Then
            MessageBox.Show(PCName.ErrorMessage, "Computer name error")
        End If
    End Sub

    Private Sub TimeZonePageTimer_Tick(sender As Object, e As EventArgs) Handles TimeZonePageTimer.Tick
        Dim UTC As Date = Date.UtcNow
        Dim SelTZ As Date = Date.UtcNow
        Dim tz As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(TimeOffsets(ComboBox5.SelectedIndex).Id)
        SelTZ = TimeZoneInfo.ConvertTimeFromUtc(SelTZ, tz)
        CurrentTimeUTC.Text = UTC.ToString("D") & " - " & UTC.ToString("HH:mm")
        CurrentTimeSelTZ.Text = SelTZ.ToString("D") & " - " & SelTZ.ToString("HH:mm")
    End Sub

    Private Sub RadioButton3_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton3.CheckedChanged
        TimeOffsetInteractive = RadioButton3.Checked
        TimeZoneSettings.Enabled = Not RadioButton3.Checked
    End Sub

    Private Sub ComboBox5_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox5.SelectedIndexChanged
        SelectedOffset = TimeOffsets(ComboBox5.SelectedIndex)
    End Sub

    Private Sub CheckBox4_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox4.CheckedChanged
        ManualPartPanel.Enabled = Not CheckBox4.Checked
        DiskConfigurationInteractive = CheckBox4.Checked
    End Sub

    Private Sub RadioButton5_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton5.CheckedChanged
        AutoDiskConfigPanel.Enabled = RadioButton5.Checked
        DiskPartPanel.Enabled = Not RadioButton5.Checked
        SelectedDiskConfiguration.DiskConfigMode = If(RadioButton5.Checked, DiskConfigurationMode.AutoDisk0, DiskConfigurationMode.DiskPart)
    End Sub

    Private Sub RadioButton7_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton7.CheckedChanged
        ESPPanel.Enabled = RadioButton7.Checked
        SelectedDiskConfiguration.PartStyle = If(RadioButton7.Checked, PartitionStyle.GPT, PartitionStyle.MBR)
    End Sub

    Private Sub CheckBox5_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox5.CheckedChanged
        WindowsREPanel.Enabled = CheckBox5.Checked
        SelectedDiskConfiguration.InstallRecEnv = CheckBox5.Checked
    End Sub

    Private Sub RadioButton9_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton9.CheckedChanged
        RESizePanel.Enabled = RadioButton9.Checked
        SelectedDiskConfiguration.RecEnvPartition = If(RadioButton9.Checked, RecoveryEnvironmentLocation.WinREPartition, RecoveryEnvironmentLocation.WindowsPartition)
    End Sub

    Private Sub RadioButton11_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton11.CheckedChanged
        ManualInstallPanel.Enabled = Not RadioButton11.Checked
        SelectedDiskConfiguration.DiskPartScriptConfig.AutomaticInstall = RadioButton11.Checked
    End Sub

    Private Sub NumericUpDown1_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown1.ValueChanged
        SelectedDiskConfiguration.ESPSize = NumericUpDown1.Value
    End Sub

    Private Sub NumericUpDown2_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown2.ValueChanged
        SelectedDiskConfiguration.RecEnvSize = NumericUpDown2.Value
    End Sub

    Private Sub NumericUpDown3_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown3.ValueChanged
        SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.DiskNum = NumericUpDown3.Value
    End Sub

    Private Sub NumericUpDown4_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown4.ValueChanged
        SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.PartNum = NumericUpDown4.Value
    End Sub

    Private Sub Scintilla2_TextChanged(sender As Object, e As EventArgs) Handles Scintilla2.TextChanged
        SelectedDiskConfiguration.DiskPartScriptConfig.ScriptContents = Scintilla2.Text
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenFileDialog1.ShowDialog()
    End Sub

    Private Sub OpenFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles OpenFileDialog1.FileOk
        Scintilla2.Text = File.ReadAllText(OpenFileDialog1.FileName)
    End Sub

    Private Sub RadioButton13_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton13.CheckedChanged
        GenericKeyPanel.Enabled = RadioButton13.Checked
        ManualKeyPanel.Enabled = Not RadioButton13.Checked
        GenericChosen = RadioButton13.Checked
    End Sub

    Private Sub ComboBox6_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox6.SelectedIndexChanged
        If GenericKeys IsNot Nothing AndAlso GenericKeys.Count > 0 Then
            SelectedKey = GenericKeys(ComboBox6.SelectedIndex)
            TextBox2.Text = SelectedKey.Key
        End If
    End Sub

    Private Sub TextBox3_TextChanged(sender As Object, e As EventArgs) Handles TextBox3.TextChanged
        SelectedKey.Key = TextBox3.Text
    End Sub

    Private Sub CheckBox6_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox6.CheckedChanged
        UserAccountsInteractive = CheckBox6.Checked
        ManualAccountPanel.Enabled = Not CheckBox6.Checked
    End Sub

#Region "User Account settings"

    Sub ModifyUserDetails(index As Integer, enabled As Boolean, name As String, password As String, group As UserGroup)
        If UserAccountsList Is Nothing OrElse UserAccountsList.Count = 0 Then Exit Sub
        UserAccountsList(index).Enabled = enabled
        UserAccountsList(index).Name = name
        UserAccountsList(index).Password = password
        UserAccountsList(index).Group = group
    End Sub

    Function GroupFromSelectedItem(index As Integer) As UserGroup
        Select Case index
            Case 0
                Return UserGroup.Administrators
            Case 1
                Return UserGroup.Users
        End Select
        Return Nothing
    End Function

    Private Sub CheckBox8_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox8.CheckedChanged
        ModifyUserDetails(1, CheckBox8.Checked, TextBox8.Text, TextBox9.Text, GroupFromSelectedItem(ComboBox9.SelectedIndex))
        TextBox8.Enabled = CheckBox8.Checked
        TextBox9.Enabled = CheckBox8.Checked
        ComboBox9.Enabled = CheckBox8.Checked
    End Sub

    Private Sub CheckBox9_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox9.CheckedChanged
        ModifyUserDetails(2, CheckBox9.Checked, TextBox11.Text, TextBox12.Text, GroupFromSelectedItem(ComboBox10.SelectedIndex))
        TextBox11.Enabled = CheckBox9.Checked
        TextBox12.Enabled = CheckBox9.Checked
        ComboBox10.Enabled = CheckBox9.Checked
    End Sub

    Private Sub CheckBox10_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox10.CheckedChanged
        ModifyUserDetails(3, CheckBox10.Checked, TextBox14.Text, TextBox15.Text, GroupFromSelectedItem(ComboBox11.SelectedIndex))
        TextBox14.Enabled = CheckBox10.Checked
        TextBox15.Enabled = CheckBox10.Checked
        ComboBox11.Enabled = CheckBox10.Checked
    End Sub

    Private Sub CheckBox11_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox11.CheckedChanged
        ModifyUserDetails(4, CheckBox11.Checked, TextBox17.Text, TextBox18.Text, GroupFromSelectedItem(ComboBox12.SelectedIndex))
        TextBox17.Enabled = CheckBox11.Checked
        TextBox18.Enabled = CheckBox11.Checked
        ComboBox12.Enabled = CheckBox11.Checked
    End Sub

    Private Sub TextBox4_TextChanged(sender As Object, e As EventArgs) Handles TextBox4.TextChanged
        ModifyUserDetails(0, True, TextBox4.Text, TextBox6.Text, GroupFromSelectedItem(ComboBox7.SelectedIndex))
    End Sub

    Private Sub TextBox6_TextChanged(sender As Object, e As EventArgs) Handles TextBox6.TextChanged
        ModifyUserDetails(0, True, TextBox4.Text, TextBox6.Text, GroupFromSelectedItem(ComboBox7.SelectedIndex))
    End Sub

    Private Sub TextBox8_TextChanged(sender As Object, e As EventArgs) Handles TextBox8.TextChanged
        ModifyUserDetails(1, CheckBox8.Checked, TextBox8.Text, TextBox9.Text, GroupFromSelectedItem(ComboBox9.SelectedIndex))
    End Sub

    Private Sub TextBox9_TextChanged(sender As Object, e As EventArgs) Handles TextBox9.TextChanged
        ModifyUserDetails(1, CheckBox8.Checked, TextBox8.Text, TextBox9.Text, GroupFromSelectedItem(ComboBox9.SelectedIndex))
    End Sub

    Private Sub TextBox11_TextChanged(sender As Object, e As EventArgs) Handles TextBox11.TextChanged
        ModifyUserDetails(2, CheckBox9.Checked, TextBox11.Text, TextBox12.Text, GroupFromSelectedItem(ComboBox10.SelectedIndex))
    End Sub

    Private Sub TextBox12_TextChanged(sender As Object, e As EventArgs) Handles TextBox12.TextChanged
        ModifyUserDetails(2, CheckBox9.Checked, TextBox11.Text, TextBox12.Text, GroupFromSelectedItem(ComboBox10.SelectedIndex))
    End Sub

    Private Sub TextBox14_TextChanged(sender As Object, e As EventArgs) Handles TextBox14.TextChanged
        ModifyUserDetails(3, CheckBox10.Checked, TextBox14.Text, TextBox15.Text, GroupFromSelectedItem(ComboBox11.SelectedIndex))
    End Sub

    Private Sub TextBox15_TextChanged(sender As Object, e As EventArgs) Handles TextBox15.TextChanged
        ModifyUserDetails(3, CheckBox10.Checked, TextBox14.Text, TextBox15.Text, GroupFromSelectedItem(ComboBox11.SelectedIndex))
    End Sub

    Private Sub TextBox17_TextChanged(sender As Object, e As EventArgs) Handles TextBox17.TextChanged
        ModifyUserDetails(4, CheckBox11.Checked, TextBox17.Text, TextBox18.Text, GroupFromSelectedItem(ComboBox12.SelectedIndex))
    End Sub

    Private Sub TextBox18_TextChanged(sender As Object, e As EventArgs) Handles TextBox18.TextChanged
        ModifyUserDetails(4, CheckBox11.Checked, TextBox17.Text, TextBox18.Text, GroupFromSelectedItem(ComboBox12.SelectedIndex))
    End Sub

    Private Sub ComboBox7_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox7.SelectedIndexChanged
        ModifyUserDetails(0, True, TextBox4.Text, TextBox6.Text, GroupFromSelectedItem(ComboBox7.SelectedIndex))
    End Sub

    Private Sub ComboBox9_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox9.SelectedIndexChanged
        ModifyUserDetails(1, CheckBox8.Checked, TextBox8.Text, TextBox9.Text, GroupFromSelectedItem(ComboBox9.SelectedIndex))
    End Sub

    Private Sub ComboBox10_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox10.SelectedIndexChanged
        ModifyUserDetails(2, CheckBox9.Checked, TextBox11.Text, TextBox12.Text, GroupFromSelectedItem(ComboBox10.SelectedIndex))
    End Sub

    Private Sub ComboBox11_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox11.SelectedIndexChanged
        ModifyUserDetails(3, CheckBox10.Checked, TextBox14.Text, TextBox15.Text, GroupFromSelectedItem(ComboBox11.SelectedIndex))
    End Sub

    Private Sub ComboBox12_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox12.SelectedIndexChanged
        ModifyUserDetails(4, CheckBox11.Checked, TextBox17.Text, TextBox18.Text, GroupFromSelectedItem(ComboBox12.SelectedIndex))
    End Sub

#End Region

    Private Sub CheckBox12_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox12.CheckedChanged
        AutoLogon.EnableAutoLogon = CheckBox12.Checked
        AutoLogonSettingsPanel.Enabled = CheckBox12.Checked
    End Sub

    Private Sub RadioButton15_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton15.CheckedChanged
        AutoLogon.LogonMode = If(RadioButton15.Checked, AutoLogonMode.FirstAdmin, AutoLogonMode.WindowsAdmin)
        TextBox5.Enabled = Not RadioButton15.Checked
    End Sub

    Private Sub TextBox5_TextChanged(sender As Object, e As EventArgs) Handles TextBox5.TextChanged
        AutoLogon.LogonPassword = TextBox5.Text
    End Sub

    Private Sub CheckBox7_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox7.CheckedChanged
        PasswordObfuscate = CheckBox7.Checked
    End Sub

    Private Sub CheckBox18_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox18.CheckedChanged
        MicrosoftAccountInteractive = CheckBox18.Checked
        AccountsPanel.Enabled = Not CheckBox18.Checked
        GroupBox1.Enabled = Not CheckBox18.Checked
        CheckBox7.Enabled = Not CheckBox18.Checked
    End Sub

    Private Sub RadioButton17_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton17.CheckedChanged
        SelectedExpirationSettings.Mode = If(RadioButton17.Checked, PasswordExpirationMode.NIST_Unlimited, PasswordExpirationMode.NIST_Limited)
        AutoExpirationPanel.Enabled = Not RadioButton17.Checked
    End Sub

    Private Sub RadioButton19_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton19.CheckedChanged
        SelectedExpirationSettings.WindowsDefault = RadioButton19.Checked
        TimedExpirationPanel.Enabled = Not RadioButton19.Checked
    End Sub

    Private Sub NumericUpDown5_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown5.ValueChanged
        SelectedExpirationSettings.Days = NumericUpDown5.Value
    End Sub

    Private Sub CheckBox13_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox13.CheckedChanged
        SelectedLockdownSettings.Enabled = CheckBox13.Checked
        EnabledAccountLockdownPanel.Enabled = Not CheckBox13.Checked
    End Sub

    Private Sub RadioButton21_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton21.CheckedChanged
        SelectedLockdownSettings.DefaultPolicy = RadioButton21.Checked
        AccountLockdownParametersPanel.Enabled = Not RadioButton21.Checked
    End Sub

    Private Sub NumericUpDown6_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown6.ValueChanged
        SelectedLockdownSettings.TimedLockdownSettings.FailedAttempts = NumericUpDown6.Value
    End Sub

    Private Sub NumericUpDown7_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown7.ValueChanged
        SelectedLockdownSettings.TimedLockdownSettings.Timeframe = NumericUpDown7.Value
    End Sub

    Private Sub NumericUpDown8_ValueChanged(sender As Object, e As EventArgs) Handles NumericUpDown8.ValueChanged
        SelectedLockdownSettings.TimedLockdownSettings.AutoUnlockTime = NumericUpDown8.Value
        NumericUpDown7.Maximum = NumericUpDown8.Value
    End Sub

    Private Sub RadioButton23_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton23.CheckedChanged
        VirtualMachineSupported = RadioButton23.Checked
        VMProviderPanel.Enabled = RadioButton23.Checked
    End Sub

    Private Sub ComboBox8_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox8.SelectedIndexChanged
        SelectedVMSettings.Provider = ComboBox8.SelectedIndex
    End Sub

    Private Sub CheckBox14_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox14.CheckedChanged
        NetworkConfigInteractive = CheckBox14.Checked
        ManualNetworkConfigPanel.Enabled = Not CheckBox14.Checked
    End Sub

    Private Sub RadioButton25_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton25.CheckedChanged
        WirelessNetworkSettingsPanel.Enabled = RadioButton25.Checked
        NetworkConfigManualSkip = Not RadioButton25.Checked
    End Sub

    Private Sub TextBox7_TextChanged(sender As Object, e As EventArgs) Handles TextBox7.TextChanged
        SelectedNetworkConfiguration.SSID = TextBox7.Text
    End Sub

    Private Sub CheckBox15_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox15.CheckedChanged
        SelectedNetworkConfiguration.ConnectWithoutBroadcast = CheckBox15.Checked
    End Sub

    Private Sub ComboBox13_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox13.SelectedIndexChanged
        SelectedNetworkConfiguration.Authentication = ComboBox13.SelectedIndex
    End Sub

    Private Sub TextBox10_TextChanged(sender As Object, e As EventArgs) Handles TextBox10.TextChanged
        SelectedNetworkConfiguration.Password = TextBox10.Text
    End Sub

    Private Sub CheckBox16_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox16.CheckedChanged
        SystemTelemetryInteractive = CheckBox16.Checked
        TelemetryOptionsPanel.Enabled = Not CheckBox16.Checked
    End Sub

    Private Sub RadioButton26_CheckedChanged(sender As Object, e As EventArgs) Handles RadioButton26.CheckedChanged
        SelectedTelemetrySettings.Enabled = Not RadioButton26.Checked
    End Sub

    Private Sub CheckBox17_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox17.CheckedChanged
        TextBox13.WordWrap = CheckBox17.Checked
    End Sub

    Private Sub LinkLabel1_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel1.LinkClicked
        Process.Start("https://schneegans.de/windows/unattend-generator/")
    End Sub

    Function EditionIDFromDisplayName(displayName As String) As String
        Select Case displayName
            Case "Home"
                Return "home"
            Case "Home N"
                Return "home_n"
            Case "Home Single Language"
                Return "home_single"
            Case "Education"
                Return "education"
            Case "Education N"
                Return "education_n"
            Case "Pro"
                Return "pro"
            Case "Pro N"
                Return "pro_n"
            Case "Pro Education"
                Return "pro_education"
            Case "Pro Education N"
                Return "pro_education_n"
            Case "Pro for Workstations"
                Return "pro_workstations"
            Case "Pro N for Workstations"
                Return "pro_workstations_n"
            Case "Enterprise"
                Return "enterprise"
        End Select
        Return ""
    End Function

    Private Sub UnattendGeneratorBW_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles UnattendGeneratorBW.DoWork
        ReportMessage("Preparing to generate file...", 0)
        If SaveTarget = "" Then
            e.Cancel = True
            Exit Sub
        End If
        ReportMessage("Preparing to generate file...", 0)
        Dim UnattendGen As New Process()
        ' Get most appropriate binary of UnattendGen
        If Environment.Is64BitOperatingSystem Then
            If PreferSelfContained Then
                UnattendGen.StartInfo.FileName = Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained\amd64\unattendgen.exe")
                UnattendGen.StartInfo.WorkingDirectory = Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained\amd64")
            Else
                UnattendGen.StartInfo.FileName = Path.Combine(Application.StartupPath, "Tools\UnattendGen\win-x64\unattendgen.exe")
                UnattendGen.StartInfo.WorkingDirectory = Path.Combine(Application.StartupPath, "Tools\UnattendGen\win-x64")
            End If
        Else
            If PreferSelfContained Then
                UnattendGen.StartInfo.FileName = Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained\x86\unattendgen.exe")
                UnattendGen.StartInfo.WorkingDirectory = Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained\x86")
            Else
                UnattendGen.StartInfo.FileName = Path.Combine(Application.StartupPath, "Tools\UnattendGen\win-x86\unattendgen.exe")
                UnattendGen.StartInfo.WorkingDirectory = Path.Combine(Application.StartupPath, "Tools\UnattendGen\win-x86")
            End If
        End If
        UnattendGen.StartInfo.Arguments = "/target=" & Quote & SaveTarget & Quote
        If Debugger.IsAttached Then
            UnattendGen.StartInfo.Arguments &= " /debug"
        End If
        Try
            ' Save settings to appropriate XML files
            ReportMessage("Saving user settings...", 2)
            Dim regSetContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                "<root>" & CrLf &
                "   <ImageLanguage Id=" & Quote & SelectedLanguage.Id & Quote & " DisplayName=" & Quote & SelectedLanguage.DisplayName & Quote & "/>" & CrLf &
                "   <UserLocale Id=" & Quote & SelectedLocale.Id & Quote & " DisplayName=" & Quote & SelectedLocale.DisplayName & Quote & " LCID=" & Quote & SelectedLocale.LCID & Quote & " KeyboardLayout=" & Quote & SelectedLocale.KeybId & Quote & " GeoLocation=" & Quote & SelectedLocale.GeoLoc & Quote & "/>" & CrLf &
                "   <KeyboardIdentifier Id=" & Quote & SelectedKeybIdentifier.Id & Quote & " DisplayName=" & Quote & SelectedKeybIdentifier.DisplayName & Quote & " Type=" & Quote & SelectedKeybIdentifier.Type & Quote & "/>" & CrLf &
                "   <GeoId Id=" & Quote & SelectedGeoId.Id & Quote & " DisplayName=" & Quote & SelectedGeoId.DisplayName & Quote & "/>" & CrLf &
                "   <TimeOffset Id=" & Quote & SelectedOffset.Id & Quote & " DisplayName=" & Quote & If(SelectedOffset.DisplayName.Contains("&"), SelectedOffset.DisplayName.Replace("&", "&amp;").Trim(), SelectedOffset.DisplayName) & Quote & "/>" & CrLf &
                "</root>"
            File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "region.xml"), regSetContents, UTF8)
            UnattendGen.StartInfo.Arguments &= " /regionfile=" & Quote & Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "region.xml") & Quote
            ReportMessage("Saving user settings...", 4)
            Select Case SelectedArchitecture
                Case DismProcessorArchitecture.Intel
                    UnattendGen.StartInfo.Arguments &= " /architecture=x86"
                Case DismProcessorArchitecture.AMD64
                    UnattendGen.StartInfo.Arguments &= " /architecture=amd64"
                Case DismProcessorArchitecture.ARM64
                    UnattendGen.StartInfo.Arguments &= " /architecture=arm64"
            End Select
            ReportMessage("Saving user settings...", 6)
            If Win11Config.LabConfig_BypassRequirements Then
                UnattendGen.StartInfo.Arguments &= " /LabConfig"
            End If
            If Win11Config.OOBE_BypassNRO Then
                UnattendGen.StartInfo.Arguments &= " /BypassNRO"
            End If
            ReportMessage("Saving user settings...", 8)
            If Not PCName.DefaultName Then
                UnattendGen.StartInfo.Arguments &= " /computername=" & PCName.Name
            End If
            ReportMessage("Saving user settings...", 10)
            If TimeOffsetInteractive Then
                UnattendGen.StartInfo.Arguments &= " /tzImplicit"
            End If
            ReportMessage("Saving user settings...", 12)
            If DiskConfigurationInteractive Then
                UnattendGen.StartInfo.Arguments &= " /partmode=interactive"
            Else
                If SelectedDiskConfiguration.DiskConfigMode = DiskConfigurationMode.AutoDisk0 Then
                    UnattendGen.StartInfo.Arguments &= " /partmode=unattended"
                    Dim diskZeroContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                        "<root>" & CrLf &
                        "   <DiskZero PartitionStyle=" & Quote & If(SelectedDiskConfiguration.PartStyle = PartitionStyle.GPT, "GPT", "MBR") & Quote & " RecoveryEnvironment=" & Quote & If(SelectedDiskConfiguration.InstallRecEnv, If(SelectedDiskConfiguration.RecEnvPartition = RecoveryEnvironmentLocation.WinREPartition, "WinRE", "Windows"), "No") & Quote & " ESPSize=" & Quote & SelectedDiskConfiguration.ESPSize & Quote & " RESize=" & Quote & SelectedDiskConfiguration.RecEnvSize & Quote & " />" & CrLf &
                        "</root>"
                    File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "unattPartSettings.xml"), diskZeroContents, UTF8)
                ElseIf SelectedDiskConfiguration.DiskConfigMode = DiskConfigurationMode.DiskPart Then
                    UnattendGen.StartInfo.Arguments &= " /partmode=custom"
                    Dim diskPartContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                        "<root>" & CrLf &
                        "   <DiskPart ScriptFile=" & Quote & Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "diskpart.dp") & Quote & " AutoInst=" & Quote & If(SelectedDiskConfiguration.DiskPartScriptConfig.AutomaticInstall, "1", "0") & Quote & " Disk=" & Quote & SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.DiskNum & Quote & " Partition=" & Quote & SelectedDiskConfiguration.DiskPartScriptConfig.TargetDisk.PartNum & Quote & " />" & CrLf &
                        "</root>"
                    File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "diskpart.dp"), SelectedDiskConfiguration.DiskPartScriptConfig.ScriptContents, UTF8)
                    File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "diskPartSettings.xml"), diskPartContents, UTF8)
                End If
            End If
            ReportMessage("Saving user settings...", 14)
            If GenericChosen Then
                UnattendGen.StartInfo.Arguments &= " /generic"
                Dim genericEditionContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                    "<root>" & CrLf &
                    "   <Edition Id=" & Quote & EditionIDFromDisplayName(ComboBox6.SelectedItem) & Quote & " DisplayName=" & Quote & ComboBox6.SelectedItem & Quote & " Key=" & Quote & SelectedKey.Key & Quote & " />" & CrLf &
                    "</root>"
                File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "edition.xml"), genericEditionContents, UTF8)
            Else
                UnattendGen.StartInfo.Arguments &= " /customkey=" & SelectedKey.Key
            End If
            If Not UserAccountsInteractive And Not MicrosoftAccountInteractive Then
                ReportMessage("Saving user settings...", 16)
                UnattendGen.StartInfo.Arguments &= " /customusers"
                Dim customUserContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                    "<root>" & CrLf
                If UserAccountsList.Count > 0 Then
                    For Each account As User In UserAccountsList
                        customUserContents &= "   <UserAccount Enabled=" & Quote & If(account.Enabled, "1", "0") & Quote & " Name=" & Quote & If(account.Name.Contains("&"), account.Name.Replace("&", "&amp;").Trim(), account.Name) & Quote & " Password=" & Quote & If(account.Password.Contains("&"), account.Password.Replace("&", "&amp;").Trim(), account.Password) & Quote & " Group=" & Quote & If(account.Group = UserGroup.Administrators, "Admins", "Users") & Quote & " />" & CrLf
                    Next
                    customUserContents &= "</root>"
                    File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "userAccounts.xml"), customUserContents, UTF8)
                    If AutoLogon.EnableAutoLogon Then
                        If AutoLogon.LogonMode = AutoLogonMode.FirstAdmin Then
                            UnattendGen.StartInfo.Arguments &= " /autologon=firstadmin"
                        ElseIf AutoLogon.LogonMode = AutoLogonMode.WindowsAdmin Then
                            UnattendGen.StartInfo.Arguments &= " /autologon=builtinadmin"
                            Dim builtinAdminContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                                "<root>" & CrLf &
                                "   <BuiltInAdmin Password=" & Quote & If(AutoLogon.LogonPassword.Contains("&"), AutoLogon.LogonPassword.Replace("&", "&amp;").Trim(), AutoLogon.LogonPassword) & Quote & " />" & CrLf &
                                "</root>"
                            File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "autoLogon.xml"), builtinAdminContents, UTF8)
                        End If
                    End If
                    If PasswordObfuscate Then
                        UnattendGen.StartInfo.Arguments &= " /b64obscure"
                    End If
                Else
                    UnattendGen.StartInfo.Arguments = UnattendGen.StartInfo.Arguments.Replace(" /customusers", "").Trim()
                End If
            ElseIf (Not UserAccountsInteractive) And MicrosoftAccountInteractive Then
                ReportMessage("Saving user settings...", 16)
                UnattendGen.StartInfo.Arguments &= " /msa"
            End If
            If SelectedExpirationSettings.Mode = PasswordExpirationMode.NIST_Limited Then
                ReportMessage("Saving user settings...", 18)
                UnattendGen.StartInfo.Arguments &= " /pwExpire=" & If(SelectedExpirationSettings.WindowsDefault, 42, SelectedExpirationSettings.Days)
            End If
            ReportMessage("Saving user settings...", 20)
            If SelectedLockdownSettings.Enabled Then
                UnattendGen.StartInfo.Arguments &= " /lockout=yes"
                Dim lockdownContents As String = ""
                If SelectedLockdownSettings.DefaultPolicy Then
                    lockdownContents = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                        "<root>" & CrLf &
                        "   <AccountLockout FailedAttempts=" & Quote & 10 & Quote & " Timeframe=" & Quote & 10 & Quote & " AutoUnlock=" & Quote & 10 & Quote & " />" & CrLf &
                        "</root>"
                Else
                    lockdownContents = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                        "<root>" & CrLf &
                        "   <AccountLockout FailedAttempts=" & Quote & SelectedLockdownSettings.TimedLockdownSettings.FailedAttempts & Quote & " Timeframe=" & Quote & SelectedLockdownSettings.TimedLockdownSettings.Timeframe & Quote & " AutoUnlock=" & Quote & SelectedLockdownSettings.TimedLockdownSettings.AutoUnlockTime & Quote & " />" & CrLf &
                        "</root>"
                End If
                File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "lockout.xml"), lockdownContents, UTF8)
            Else
                UnattendGen.StartInfo.Arguments &= " /lockout=no"
            End If
            If VirtualMachineSupported Then
                ReportMessage("Saving user settings...", 22)
                Select Case SelectedVMSettings.Provider
                    Case VMProvider.VirtualBox_GAs
                        UnattendGen.StartInfo.Arguments &= " /vm=vbox_gas"
                    Case VMProvider.VMware_Tools
                        UnattendGen.StartInfo.Arguments &= " /vm=vmware"
                    Case VMProvider.VirtIO_Guest_Tools
                        UnattendGen.StartInfo.Arguments &= " /vm=virtio"
                End Select
            End If
            If Not NetworkConfigInteractive Then
                ReportMessage("Saving user settings...", 24)
                If NetworkConfigManualSkip Then
                    UnattendGen.StartInfo.Arguments &= " /wifi=no"
                Else
                    UnattendGen.StartInfo.Arguments &= " /wifi=yes"
                    Dim wirelessContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                        "<root>" & CrLf &
                        "   <WirelessNetwork Name=" & Quote & If(SelectedNetworkConfiguration.SSID.Contains("&"), SelectedNetworkConfiguration.SSID.Replace("&", "&amp;").Trim(), SelectedNetworkConfiguration.SSID) & Quote & " Password=" & Quote & If(SelectedNetworkConfiguration.Password.Contains("&"), SelectedNetworkConfiguration.Password.Replace("&", "&amp;").Trim(), SelectedNetworkConfiguration.Password) & Quote & " AuthMode=" & Quote & If(SelectedNetworkConfiguration.Authentication = WiFiAuthenticationMode.Open, "Open", If(SelectedNetworkConfiguration.Authentication = WiFiAuthenticationMode.WPA2_PSK, "WPA2", "WPA3")) & Quote & " NonBroadcast=" & Quote & If(SelectedNetworkConfiguration.ConnectWithoutBroadcast, "1", "0") & Quote & " />" & CrLf &
                        "</root>"
                    File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "wireless.xml"), wirelessContents, UTF8)
                End If
            End If
            If Not SystemTelemetryInteractive Then
                ReportMessage("Saving user settings...", 24.5)
                If SelectedTelemetrySettings.Enabled Then
                    UnattendGen.StartInfo.Arguments &= " /telem=yes"
                Else
                    UnattendGen.StartInfo.Arguments &= " /telem=no"
                End If
            End If
            If FinalComponents.Count > 0 Then
                ReportMessage("Saving user settings...", 24.75)
                UnattendGen.StartInfo.Arguments &= " /customcomponents"
                Dim customComponentContents As String = "<?xml version=" & Quote & "1.0" & Quote & " ?>" & CrLf &
                    "<root>" & CrLf
                For Each systemComponent As Component In FinalComponents
                    Dim passName As String = ""
                    If systemComponent.Passes.Count > 0 Then
                        For Each systemPass As Pass In systemComponent.Passes
                            passName &= systemPass.Name & ","
                        Next
                        passName = passName.TrimEnd(",")
                    End If
                    customComponentContents &= "    <Component Id=" & Quote & systemComponent.Id.Replace("&", "&amp;").Trim() & Quote & " Passes=" & Quote & passName.Replace("&", "&amp;").Trim() & Quote & " />" & CrLf
                Next
                customComponentContents &= "</root>"
                File.WriteAllText(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "components.xml"), customComponentContents, UTF8)
            End If
            ReportMessage("Generating unattended answer file...", 25)
            UnattendGen.Start()
            UnattendGen.WaitForExit()
            ReportMessage("Generating unattended answer file...", 50)
            ReportMessage("Deleting temporary files...", 75)
            If File.Exists(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "diskpart.dp")) Then
                File.Delete(Path.Combine(UnattendGen.StartInfo.WorkingDirectory, "diskpart.dp"))
            End If
            For Each xmlFile In My.Computer.FileSystem.GetFiles(UnattendGen.StartInfo.WorkingDirectory, FileIO.SearchOption.SearchTopLevelOnly, "*.xml")
                If File.Exists(xmlFile) Then File.Delete(xmlFile)
            Next
            If UnattendGen.ExitCode <> 0 Then
                MessageBox.Show("The unattended answer file generator could not generate the file. Here is the error code if you are interested" & CrLf & CrLf & "Error code: " & Hex(UnattendGen.ExitCode))
                e.Cancel = True
            End If
            ReportMessage("Generation has completed", 100)
        Catch ex As Exception
            If UnattendGen.ExitCode <> 0 Then
                MessageBox.Show("The unattended answer file generator could not generate the file. Here is the error code if you are interested" & CrLf & CrLf & "Error: " & ex.Message)
                e.Cancel = True
            End If
        End Try
    End Sub

    Private Sub SaveFileDialog1_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles SaveFileDialog1.FileOk
        SaveTarget = SaveFileDialog1.FileName
    End Sub

    Sub ReportMessage(msg As String, percent As Integer)
        ProgressMessage = msg
        UnattendGeneratorBW.ReportProgress(percent)
    End Sub

    Private Sub UnattendGeneratorBW_ProgressChanged(sender As Object, e As System.ComponentModel.ProgressChangedEventArgs) Handles UnattendGeneratorBW.ProgressChanged
        Label56.Text = ProgressMessage
        ProgressBar1.Value = e.ProgressPercentage
    End Sub

    Private Sub UnattendGeneratorBW_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles UnattendGeneratorBW.RunWorkerCompleted
        If e.Cancelled Then
            ChangePage(CurrentWizardPage.WizardPage - 1)
            Exit Sub
        End If
        ChangePage(CurrentWizardPage.WizardPage + 1)
    End Sub

    Private Sub LinkLabel2_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel2.LinkClicked
        If MsgBox("Do you want to reuse the settings you've used in this answer file for the new one?", vbQuestion + vbYesNo, Text) = MsgBoxResult.No Then
            ' Refresh the settings
            ReloadSettings()
        End If
        ChangePage(UnattendedWizardPage.Page.RegionalPage)
    End Sub

    Private Sub LinkLabel3_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel3.LinkClicked
        Process.Start(Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\explorer.exe", "/select," & Quote & SaveTarget & Quote)
    End Sub

    Private Sub LinkLabel4_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel4.LinkClicked
        If MainForm.isProjectLoaded And Not (MainForm.OnlineManagement Or MainForm.OfflineManagement) Then
            ApplyUnattendFile.TextBox1.Text = SaveTarget
            WindowState = FormWindowState.Minimized
            ApplyUnattendFile.ShowDialog(MainForm)
            WindowState = FormWindowState.Normal
        Else
            MsgBox("You need to load a project in order to apply this file.", vbOKOnly + vbExclamation, Text)
            Exit Sub
        End If
    End Sub

    Private Sub NewUnattendWiz_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        If UnattendGeneratorBW.IsBusy OrElse UnattendGenBW.IsBusy Then
            e.Cancel = True
            Beep()
            Exit Sub
        End If
    End Sub

    Private Sub StepsTreeView_DrawNode(sender As Object, e As DrawTreeNodeEventArgs) Handles StepsTreeView.DrawNode
        ' Determine the custom background color
        Dim customBackColor As Color = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.FromArgb(31, 31, 31), Color.FromArgb(239, 239, 242))

        ' Determine the custom foreground color based on the custom background color
        Dim customForeColor As Color = If(customBackColor = Color.FromArgb(31, 31, 31), Color.White, Color.Black)

        ' Check if the node is selected
        If (e.State And TreeNodeStates.Selected) <> 0 Then
            ' Draw the background with the highlight color
            e.Graphics.FillRectangle(SystemBrushes.Highlight, e.Bounds)
            ' Draw the text with the highlighted text color and vertically centered
            TextRenderer.DrawText(e.Graphics, e.Node.Text, If(e.Node.NodeFont, e.Node.TreeView.Font), e.Bounds, SystemColors.HighlightText, TextFormatFlags.VerticalCenter)
        Else
            ' Draw the background with the custom color for unselected nodes
            Using backgroundBrush As New SolidBrush(customBackColor)
                e.Graphics.FillRectangle(backgroundBrush, e.Bounds)
            End Using
            ' Draw the text with the custom foreground color and vertically centered
            TextRenderer.DrawText(e.Graphics, e.Node.Text, If(e.Node.NodeFont, e.Node.TreeView.Font), e.Bounds, customForeColor, TextFormatFlags.VerticalCenter)
        End If

        ' If the node has focus, draw the focus rectangle
        If (e.State And TreeNodeStates.Focused) <> 0 Then
            Using focusPen As New Pen(Color.Black)
                focusPen.DashStyle = Drawing2D.DashStyle.Dot
                Dim focusBounds As Rectangle = e.Bounds
                focusBounds.Size = New Size(focusBounds.Width - 1, focusBounds.Height - 1)
                e.Graphics.DrawRectangle(focusPen, focusBounds)
            End Using
        End If

        ' Signal that the node has been drawn
        e.DrawDefault = False
    End Sub

    Private Sub UnattendGenBW_DoWork(sender As Object, e As System.ComponentModel.DoWorkEventArgs) Handles UnattendGenBW.DoWork
        Try
            ' Download UnattendGen and run it
            If Not Directory.Exists(Application.StartupPath & "\Tools\UnattendGen\SelfContained") Then
                Directory.CreateDirectory(Application.StartupPath & "\Tools\UnattendGen\SelfContained")
            End If
            Using UnattClient As New WebClient()
                ServicePointManager.SecurityProtocol = SecurityProtocolType.Tls12
                Dim contents As String = ""
                Try
                    contents = UnattClient.DownloadString("https://raw.githubusercontent.com/CodingWonders/UnattendGen/master/DISMTools-Install.ps1")
                Catch ex As WebException
                    Throw ex
                End Try
                If contents <> "" Then
                    File.WriteAllText(Application.StartupPath & "\setup.ps1", contents, UTF8)
                End If
            End Using
            If File.Exists(Application.StartupPath & "\setup.ps1") Then
                ' Run installer
                Dim UAProc As New Process()
                UAProc.StartInfo.FileName = Environment.GetFolderPath(Environment.SpecialFolder.Windows) & "\system32\WindowsPowerShell\v1.0\powershell.exe"
                UAProc.StartInfo.WorkingDirectory = Application.StartupPath
                UAProc.StartInfo.Arguments = "-executionpolicy unrestricted -file " & Quote & Application.StartupPath & "\setup.ps1" & Quote & " -tag " & Quote & "DT_" & UnattendGenReleaseTag & Quote
                UAProc.Start()
                UAProc.WaitForExit()
                If UAProc.ExitCode <> 0 Then
                    Throw New System.ComponentModel.Win32Exception(UAProc.ExitCode)
                End If
            End If
            If File.Exists(Application.StartupPath & "\setup.ps1") Then
                Try
                    File.Delete(Application.StartupPath & "\setup.ps1")
                Catch ex As Exception
                    ' Don't delete it
                End Try
            End If
        Catch ex As Exception
            Throw ex
        End Try
    End Sub

    Private Sub UnattendGenBW_RunWorkerCompleted(sender As Object, e As System.ComponentModel.RunWorkerCompletedEventArgs) Handles UnattendGenBW.RunWorkerCompleted
        If e.Error IsNot Nothing Then
            MessageBox.Show("We couldn't prepare UnattendGen Self-Contained Setup. Reason:" & CrLf & e.Error.Message, "UnattendGen error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            If Directory.Exists(Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained")) Then
                Try
                    Directory.Delete(Path.Combine(Application.StartupPath, "Tools\UnattendGen\SelfContained"), True)
                Catch ex As Exception
                    ' Leave dir
                End Try
            End If
            Close()
            Exit Sub
        End If
        ExpressPanelFooter.Enabled = True
        PreferSelfContained = True
        UGNotify.ShowBalloonTip(5000)
    End Sub

    Private Sub ToolStripButton2_Click(sender As Object, e As EventArgs) Handles ToolStripButton2.Click
        Scintilla1.Text = DefaultContents
    End Sub

    Private Sub ToolStripButton3_Click(sender As Object, e As EventArgs) Handles ToolStripButton3.Click
        EditorModeOFD.ShowDialog()
    End Sub

    Private Sub EditorModeOFD_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles EditorModeOFD.FileOk
        Try
            Scintilla1.Text = File.ReadAllText(EditorModeOFD.FileName)
        Catch ex As Exception
            MsgBox("Could not open file: " & ex.Message, vbOKOnly + vbCritical, Text)
        End Try
    End Sub

    Private Sub ToolStripButton4_Click(sender As Object, e As EventArgs) Handles ToolStripButton4.Click
        EditorModeSFD.ShowDialog()
    End Sub

    Private Sub EditorModeSFD_FileOk(sender As Object, e As System.ComponentModel.CancelEventArgs) Handles EditorModeSFD.FileOk
        Try
            File.WriteAllText(EditorModeSFD.FileName, Scintilla1.Text, UTF8)
        Catch ex As Exception
            MsgBox("Could not save file: " & ex.Message, vbOKOnly + vbCritical, Text)
        End Try
    End Sub

    Private Sub Help_Button_Click(sender As Object, e As EventArgs) Handles Help_Button.Click, ToolStripButton6.Click
        HelpBrowserForm.WebBrowser1.Navigate(Application.StartupPath & "\docs\img_tasks\unattend\unatt_create\index.html")
        HelpBrowserForm.MinimizeBox = False
        HelpBrowserForm.MaximizeBox = False
        HelpBrowserForm.ShowDialog()
    End Sub

    Private Sub NewUnattendWiz_SizeChanged(sender As Object, e As EventArgs) Handles MyBase.SizeChanged
        AutoDiskConfigPanel.Width = ManualPartPanel.Width - (AutoDiskConfigPanel.Margin.Left * 2) - 4
        DiskPartPanel.Width = ManualPartPanel.Width - (DiskPartPanel.Margin.Left * 2) - 4
        GroupBox1.Width = ManualAccountPanel.Width - (GroupBox1.Margin.Left * 2) - 4
        AccountsPanel.Width = UserAccountListing.Width
        UserAccountListing.Width = ManualAccountPanel.Width - (UserAccountListing.Margin.Left * 2) - 4
        WirelessNetworkSettingsPanel.Width = ManualNetworkConfigPanel.Width - (WirelessNetworkSettingsPanel.Margin.Left * 2) - 4
    End Sub

    Private Sub ListBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ListBox2.SelectedIndexChanged
        PassConfigurationPanel.Visible = (ListBox2.SelectedItems.Count = 1)
        If ListBox2.SelectedItems.Count = 1 Then
            For Each configurationPass As Pass In SystemComponents(ListBox2.SelectedIndex).Passes
                Select Case configurationPass.Name
                    Case "windowsPE"
                        windowsPE.Enabled = configurationPass.Compatible
                        windowsPE.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "offlineServicing"
                        offlineServicing.Enabled = configurationPass.Compatible
                        offlineServicing.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "specialize"
                        specialize.Enabled = configurationPass.Compatible
                        specialize.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "generalize"
                        generalize.Enabled = configurationPass.Compatible
                        generalize.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "auditSystem"
                        auditSystem.Enabled = configurationPass.Compatible
                        auditSystem.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "auditUser"
                        auditUser.Enabled = configurationPass.Compatible
                        auditUser.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                    Case "oobeSystem"
                        oobeSystem.Enabled = configurationPass.Compatible
                        oobeSystem.Checked = If(configurationPass.Compatible, configurationPass.Enabled, False)
                End Select
            Next
        End If
    End Sub

    Private Sub LinkLabel5_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel5.LinkClicked
        Process.Start("https://learn.microsoft.com/en-us/windows-hardware/customize/desktop/unattend/components-b-unattend")
    End Sub

    Sub ConfigureComponent(componentName As String, componentPass As String, componentPassEnabled As Boolean)
        If String.IsNullOrWhiteSpace(componentName) Then Exit Sub
        If String.IsNullOrWhiteSpace(componentPass) Then Exit Sub
        Dim componentNames As New List(Of String)
        Dim knownPasses As New Dictionary(Of String, Boolean)
        knownPasses.Add("offlineServicing", False)
        knownPasses.Add("windowsPE", False)
        knownPasses.Add("generalize", False)
        knownPasses.Add("specialize", False)
        knownPasses.Add("auditSystem", False)
        knownPasses.Add("auditUser", False)
        knownPasses.Add("oobeSystem", False)
        For Each systemComponent As Component In SystemComponents
            componentNames.Add(systemComponent.Id)
        Next
        ' Determine if the passed component ID "componentName" exists in the grabbed components
        If componentNames.Contains(componentName) Then
            Dim placementIndex As Integer = componentNames.IndexOf(componentName)
            ' Grab pass to configure and configure it
            If Not knownPasses.ContainsKey(componentPass) Then
                MsgBox("The component pass " & componentPass & " does not exist in the pass list", vbOKOnly + vbCritical, Text)
                Exit Sub
            End If
            Dim editedPass As Pass = SystemComponents(placementIndex).Passes.FirstOrDefault(Function(p) p.Name = componentPass)
            If editedPass IsNot Nothing Then
                editedPass.Enabled = componentPassEnabled
                Debug.WriteLine("The pass " & Quote & componentPass & Quote & " of the component " & Quote & SystemComponents(placementIndex).Id & Quote & " has been " & If(SystemComponents(placementIndex).Passes.FirstOrDefault(Function(p) p.Name = componentPass).Enabled, "enabled", "disabled"))
            Else

            End If
        Else
            MsgBox("The component " & componentName & " does not exist in the component list", vbOKOnly + vbCritical, Text)
            Exit Sub
        End If
    End Sub

    Private Sub oobeSystem_CheckedChanged(sender As Object, e As EventArgs) Handles oobeSystem.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "oobeSystem", oobeSystem.Checked)
    End Sub

    Private Sub auditUser_CheckedChanged(sender As Object, e As EventArgs) Handles auditUser.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "auditUser", auditUser.Checked)
    End Sub

    Private Sub auditSystem_CheckedChanged(sender As Object, e As EventArgs) Handles auditSystem.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "auditSystem", auditSystem.Checked)
    End Sub

    Private Sub generalize_CheckedChanged(sender As Object, e As EventArgs) Handles generalize.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "generalize", generalize.Checked)
    End Sub

    Private Sub specialize_CheckedChanged(sender As Object, e As EventArgs) Handles specialize.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "specialize", specialize.Checked)
    End Sub

    Private Sub offlineServicing_CheckedChanged(sender As Object, e As EventArgs) Handles offlineServicing.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "offlineServicing", offlineServicing.Checked)
    End Sub

    Private Sub windowsPE_CheckedChanged(sender As Object, e As EventArgs) Handles windowsPE.CheckedChanged
        ConfigureComponent(ListBox2.SelectedItem, "windowsPE", windowsPE.Checked)
    End Sub

    Function AreComponentListsEqual(list1 As List(Of Component), list2 As List(Of Component)) As Boolean
        ' Check if the counts of both lists are the same
        If list1.Count <> list2.Count Then Return False

        ' Iterate through components in both lists
        For i As Integer = 0 To list1.Count - 1
            Dim component1 As Component = list1(i)
            Dim component2 As Component = list2(i)

            ' Compare component IDs
            If component1.Id <> component2.Id Then Return False

            ' Compare the number of passes in each component
            If component1.Passes.Count <> component2.Passes.Count Then Return False

            ' Compare each pass
            For j As Integer = 0 To component1.Passes.Count - 1
                Dim pass1 As Pass = component1.Passes(j)
                Dim pass2 As Pass = component2.Passes(j)

                ' Compare pass names and compatible states
                If pass1.Name <> pass2.Name OrElse pass1.Enabled <> pass2.Enabled Then
                    Return False
                End If
            Next
        Next

        ' If all comparisons pass, the lists are equal
        Return True
    End Function

    Function GetComponentDifferences(list1 As List(Of Component), list2 As List(Of Component)) As List(Of Component)
        Dim differences As New List(Of Component)

        ' Combine both lists to check for differences in either
        Dim allComponents As List(Of Component) = list1.Concat(list2).GroupBy(Function(c) c.Id).Select(Function(g) g.First()).ToList()

        For Each component In allComponents
            ' Find the component in both lists
            Dim component1 As Component = list1.FirstOrDefault(Function(c) c.Id = component.Id)
            Dim component2 As Component = list2.FirstOrDefault(Function(c) c.Id = component.Id)

            If component1 Is Nothing OrElse component2 Is Nothing Then
                ' If a component is missing in one of the lists, it's different
                differences.Add(component)
            Else
                ' Compare passes if the component exists in both lists
                Dim differingComponent As New Component() With {.Id = component.Id}
                For Each pass1 In component1.Passes
                    ' Find corresponding pass in component2
                    Dim pass2 As Pass = component2.Passes.FirstOrDefault(Function(p) p.Name = pass1.Name)

                    If pass2 Is Nothing OrElse pass1.Enabled <> pass2.Enabled Then
                        ' If pass is missing or its status is different, mark it as different
                        differingComponent.Passes.Add(pass1)
                    End If
                Next

                ' Only add the component if there are differing passes
                If differingComponent.Passes.Count > 0 Then
                    differences.Add(differingComponent)
                End If
            End If
        Next

        Return differences
    End Function

    Private Sub LinkLabel6_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel6.LinkClicked
        If File.Exists(Path.Combine(Environment.GetFolderPath(If(Environment.Is64BitOperatingSystem, Environment.SpecialFolder.ProgramFilesX86, Environment.SpecialFolder.ProgramFiles)),
                                    "Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\WSIM\x86\imgmgr.exe")) Then
            Process.Start(Path.Combine(Environment.GetFolderPath(If(Environment.Is64BitOperatingSystem, Environment.SpecialFolder.ProgramFilesX86, Environment.SpecialFolder.ProgramFiles)), "Windows Kits\10\Assessment and Deployment Kit\Deployment Tools\WSIM\x86\imgmgr.exe"), Quote & SaveTarget & Quote)
        End If
    End Sub

    Private Sub LinkLabel7_LinkClicked(sender As Object, e As LinkLabelLinkClickedEventArgs) Handles LinkLabel7.LinkClicked
        Try
            Scintilla1.Text = File.ReadAllText(SaveTarget)
        Catch ex As Exception
            MsgBox("Could not open file: " & ex.Message, vbOKOnly + vbCritical, Text)
            Exit Sub
        End Try

        IsInExpress = False
        StepsTreeView.Enabled = False
        EditorPanelContainer.Visible = True
        ExpressPanelContainer.Visible = False
        ExpressPanelTrigger.BackColor = SidePanel.BackColor
        ExpressPanelTrigger.ForeColor = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), Color.LightGray, Color.Black)
        PictureBox1.Image = If(MainForm.BackColor = Color.FromArgb(48, 48, 48), My.Resources.express_mode_select, My.Resources.express_mode)
        EditorPanelTrigger.BackColor = Color.FromKnownColor(KnownColor.Highlight)
        EditorPanelTrigger.ForeColor = Color.White
        PictureBox2.Image = My.Resources.editor_mode_select
        PictureBox3.Image = My.Resources.editor_mode_fc
        Label3.Text = "Editor mode"
        Label4.Text = "Create your unattended answer files from scratch and save them anywhere"
        FooterContainer.Visible = False
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        TextBox1.Text = My.Computer.Name
    End Sub

    Private Sub Button3_MouseHover(sender As Object, e As EventArgs) Handles Button3.MouseHover
        CNameTTip.Show("Uses the name of your computer as the computer name of the unattended answer file." & CrLf & "Only use this if the system you want to target is this one", sender)
    End Sub
End Class