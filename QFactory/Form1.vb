Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Imports iNovation.Code.Encryption
Imports iNovation.Code.Media
Imports iNovation.Code.Sequel
Imports iNovation.Code.Machine
'Imports iNovation.Code.Web

Public Class Form1
#Region "LanguageC"
    Private g As New iNovation.Code.Desktop '' General_Module.FormatWindow()
    Public f As New iNovation.Code.Feedback '' Feedback.Feedback
    ''    Public l As New Language.Languages



#Region "File Paths Variables"

    'Public ReadOnly Property notification_file As String
    '	Get
    '		Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & app & "\Settings_Notification.txt"
    '	End Get
    'End Property

    'Public Shared ReadOnly Property logo As String
    '	Get
    '		Return Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & app & "\logo.jpg"
    '	End Get
    'End Property


    'Public Shared ReadOnly Property starter__ As String
    '	Get
    '		Return My.Application.Info.DirectoryPath & "\My Admin Starter.exe"
    '	End Get
    'End Property

    Public ReadOnly Property printer_ As String
        Get
            Return My.Settings.printer_ ' Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\" & app & "\Setting_Printer.txt"
        End Get
    End Property

#End Region

#End Region


#Region "Database"
    Public ReadOnly font_alt As Font = New Font("Verdana", 12, GraphicsUnit.Point)
    Public ReadOnly font_brand As Font = New Font("Miriam Fixed", 26, FontStyle.Bold, GraphicsUnit.Point)

    Public ReadOnly font_ As Font = New Font("Verdana", 10, GraphicsUnit.Point)
    Public ReadOnly theme_file As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\iNovation Digital Works\QFactory\Theme.txt"
    Public ReadOnly theme_background_file As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\iNovation Digital Works\QFactory\ThemeBackground.txt"
    Public ReadOnly source_file As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\iNovation Digital Works\QFactory\Source.txt"
    Public ReadOnly top_rows_file As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\inovation digital works\QFactory\TopRows.txt"


#End Region

#Region "In/Out Timer"
    Private Sub CloseButton_Click(sender As Object, e As EventArgs) Handles CloseButton.Click
        '        Bye(Me, "Take care", OutTimer, False)
        ''        f.say("Exiting Factory")
        ''       My.Computer.Audio.Play(log_out_sound, AudioPlayMode.WaitToComplete)
        ''       OutTimer.Enabled = True
    End Sub
    'Private server_is_checked As Boolean = False
    Private Sub InTimer_Tick(sender As Object, e As EventArgs) Handles InTimer.Tick
        If Me.Opacity >= 1 Then
            InTimer.Enabled = False
            Exit Sub
        End If
        Me.Opacity += 0.2
    End Sub

    Private Sub OutTimer_Tick(sender As Object, e As EventArgs) Handles OutTimer.Tick
        If Me.Opacity <= 0 Then
            Environment.Exit(0)
        End If
        Me.Opacity -= 0.2
    End Sub

#End Region

#Region "theme"
    Private Sub GreenTheme_Click(sender As Object, e As EventArgs) Handles GreenTheme.Click, TurqoiseTheme.Click, BrownTheme.Click, YellowTheme.Click, VelvetTheme.Click, PurpleTheme.Click, WhiteTheme.Click
        Dim mark_ As ToolStripMenuItem = sender
        'MarkTheme(mark_.Text, GreenTheme, TurqoiseTheme, VelvetTheme, PurpleTheme, WhiteTheme, BrownTheme, YellowTheme, Me)
        'g.item_f = font_
        'g.FormatMe(Me, Nothing, LeftBorder, RightBorder, TopBorder, BottomBorder, TopLine, BottomLine, DialogTitle, SystemCloseButton, MinimizeButton, CloseButton, HelpButton, MenuStrip, SystemCloseButton, SystemCloseButton, TitleBar, FooterBar, MaximizeButton, True, False, False, True, selected_theme, Feedback, TimeTimer, TimeLabel, False, False)
        TopLine.Hide()
        BottomLine.Hide()
        DialogTitle.Hide()
        gColumns.ColumnHeadersHeight = 24
        grid.ColumnHeadersHeight = 24

        With tOperators
            .BorderStyle = BorderStyle.None
            .ScrollBars = ScrollBars.None
        End With
    End Sub

    Private Sub ThemeSelectBackground_Click(sender As Object, e As EventArgs) Handles ThemeSelectBackground.Click
        'DialogBackground(Me)
    End Sub

    Private Sub ThemeClearBackground_Click(sender As Object, e As EventArgs) Handles ThemeClearBackground.Click
        'ClearDialogBackground(Me)
    End Sub
#End Region


#Region "Functions"

    Private Enum Query
        BuildCountString
        BuildCountString_CONDITIONAL
        BuildInsertString
        BuildMaxString
        BuildSelectString
        BuildSelectString_BETWEEN
        BuildSelectString_CONDITIONAL
        BuildSelectString_DISTINCT
        BuildSelectString_LIKE
        BuildTopString
        BuildUpdateString
        BuildUpdateString_CONDITIONAL
        DeleteString_CONDITIONAL
        BuildTopString_CONDITIONAL
    End Enum

    Public Sub GetColumns(table_ As String)
        Clear(dColumns)
        Clear(lColumns)
        Clear(dWhere)
        Clear(lWhere)
        '        Clear(dQueryOperators)
        Clear(lQueryOperators)
        Clear(dMax)
        Dim l As New List(Of String)
        Try
            Display(gColumns, BuildSelectString(table_), tString.Text)
        Catch ex As ArgumentException
        End Try
        With gColumns
            'For i As Integer = 0 To .Columns.Count - 1
            '    If .Columns.Item(i).HeaderText.ToLower.Contains("picture") Or .Columns.Item(i).HeaderText.ToLower.Contains("logo") Then
            '        .Columns.Item(i).Visible = False
            '    End If
            'Next

            For i As Integer = 0 To .Columns.Count - 1
                '                If i > 0 Then Exit For
                l.Add(.Columns.Item(i).HeaderText)
                dColumns.Items.Add(.Columns.Item(i).HeaderText)
                dWhere.Items.Add(.Columns.Item(i).HeaderText)
                dMax.Items.Add(.Columns.Item(i).HeaderText)
            Next
            dMax.Text = .Columns.Item(0).HeaderText
        End With
        'With l
        '    For j As Integer = 0 To .Count - 1
        '        dColumns.Items.Add(l(j))
        '        dWhere.Items.Add(l(j))
        '    Next
        'End With
    End Sub
    Public Sub GetTables() 'group_ As Group
        Dim q__ As String = BuildSelectString("sys.Tables", {"name"}, Nothing, "name")
        'Display(grid, q__, Content(tString))
        Clear(dTable)
        Try
            ''    Clear(dTable)
            ''    Select Case Content(dGroup) ' group_
            '        Case "All"
            'BindProperty(dTable, "", q__, Content(tString), Nothing, "name")
            With grid
                For i As Integer = 0 To .Rows.Count - 1
                    dTable.Items.Add(.Rows(i).Cells(0).Value.ToString)
                Next
            End With
            'Case "Other"
            '    With grid
            '        For i As Integer = 0 To .Rows.Count - 1
            '            If .Rows(i).Cells(0).Value.ToString.Contains("MyAdmin") = False Then
            '                dTable.Items.Add(.Rows(i).Cells(0).Value.ToString)
            '            End If
            '        Next
            '    End With
            'Case "My Admin" ' Group.MyAdmin.ToString.Replace("MyAdmin", "My Admin")
            '    'q__ = BuildSelectString_LIKE("sys.Tables", {"name"}, {"name"})
            '    'BindProperty(dTable, "", q__, Content(tString), {"name", "MyAdmin"}, "name")
            '    With grid
            '        For i As Integer = 0 To .Rows.Count - 1
            '            If .Rows(i).Cells(0).Value.ToString.Contains("MyAdmin") Then
            '                dTable.Items.Add(.Rows(i).Cells(0).Value.ToString)
            '            End If
            '        Next
            '    End With
            'End Select
        Catch
        End Try
    End Sub
    Public Sub GetDataSource()
        tString.Text = If(IsEmpty(textDataSource), "Data Source=" & Environment.MachineName & "\SQLEXPRESS;Initial Catalog=" & Content(dCatalog) & ";Persist Security Info=True;User ID=" & Content(textUserID) & ";Password=" & Content(textPassword), "Data Source=" & Content(textDataSource) & "\SQLEXPRESS;Initial Catalog=" & Content(dCatalog) & ";Persist Security Info=True;User ID=" & Content(textUserID) & ";Password=" & Content(textPassword))

        Clear(dTable)
        dQueryFromModule.Text = ""
        Clear(dColumns)
        Clear(dWhere)
        Clear(lColumns)
        Clear(lWhere)
        Clear(lQueryOperators)
        Try
            Dim q__ As String = BuildSelectString("sys.Tables", {"name"}, Nothing, "name")
            Display(grid, q__, Content(tString))
        Catch ex As Exception
        End Try
    End Sub
    'Public Sub ClearFields()
    '    GetDataSource()

    '    Try
    '        GetTables() '(Content(dGroup))
    '    Catch
    '    End Try
    'End Sub
    Public Sub SaveProgramSettings()
        'Return
        '        My.Settings.source__ = Content(dSource)
        'WriteText(source_file, Content(dSource))
    End Sub

    Private Function Prepend(query_ As String) As String
        Dim return_ As String = "USE " & Content(dCatalog) & vbCrLf '& vbCrLf
        Dim return__ As String
        '        If hSQL.Checked = False Then
        '       return__ = query_
        '      Else
        '       If Content(dQueryFromModule).ToLower.Contains("insert") Or Content(dQueryFromModule).ToLower.Contains("update") Then
        Try
            With lColumns
                .SelectedIndex = 0
                For i As Integer = 0 To .Items.Count - 1
                    return_ &= "DECLARE @" & .SelectedItem & " nvarchar(max)"
                    return_ &= vbCrLf & "SET @" & .SelectedItem & "=''"
                    return_ &= vbCrLf
                    Try
                        .SelectedIndex += 1
                    Catch ex As Exception

                    End Try
                Next
            End With
        Catch
        End Try
        return_ &= vbCrLf
        '        End If
        Try
            With lWhere
                .SelectedIndex = 0
                For i As Integer = 0 To .Items.Count - 1
                    If query_.ToLower.Contains("between") Then
                        return_ &= "DECLARE @" & .SelectedItem & "_FROM nvarchar(max)"
                        return_ &= vbCrLf & "SET @" & .SelectedItem & "_FROM=''"
                        return_ &= vbCrLf
                        return_ &= "DECLARE @" & .SelectedItem & "_TO nvarchar(max)"
                        return_ &= vbCrLf & "SET @" & .SelectedItem & "_TO=''"
                        return_ &= vbCrLf
                    Else
                        return_ &= "DECLARE @" & .SelectedItem & " nvarchar(max)"
                        return_ &= vbCrLf & "SET @" & .SelectedItem & "=''"
                        return_ &= vbCrLf
                    End If
                    Try
                        .SelectedIndex += 1
                    Catch ex As Exception

                    End Try
                Next
            End With
        Catch
        End Try
        return__ = return_ & vbCrLf & vbCrLf & query_
        '        End If
        Return return__.Trim

    End Function

    Private Function Code() As String
        If Content(dQueryFromModule).ToLower.Contains("delete") And Content(dOutput).ToLower <> "qstring" Then Return ""

        If Content(dOutput).ToLower = "commit" Then
            If Content(dQueryFromModule).ToLower.Contains("update") = False And Content(dQueryFromModule).ToLower.Contains("insert") = False Then Return ""
        End If

        If Content(dOutput).ToLower = "display" Then
            If Content(dQueryFromModule).ToLower.Contains("update") Or Content(dQueryFromModule).ToLower.Contains("insert") Then Return ""
        End If

        If Content(dQueryFromModule).ToLower.Contains("update") Or Content(dQueryFromModule).ToLower.Contains("insert") Then
            If Content(dOutput).ToLower <> "commit" Then Return ""
        End If

        Dim r As String = ""
        Dim q_params As QParameters
        With q_params
            .InsertKeys = ListToArray(lColumns, ReturnInfo.AsArray)
            Select Case Content(dLikeSelect).ToLower
                Case "and"
                    .LikeSelect = LIKE_SELECT.AND_
                Case "or"
                    .LikeSelect = LIKE_SELECT.OR_
            End Select
            .MaxField = Content(dMax)
            Select Case Content(dQueryFromModule).ToLower
                Case "BuildCountString".ToLower
                    .operation = Queries.BuildCountString

                Case "BuildCountString_CONDITIONAL".ToLower
                    .operation = Queries.BuildCountString_CONDITIONAL

                Case "BuildInsertString".ToLower
                    .operation = Queries.BuildInsertString

                Case "BuildMaxString".ToLower
                    .operation = Queries.BuildMaxString

                Case "BuildSelectString".ToLower
                    .operation = Queries.BuildSelectString

                Case "BuildSelectString_CONDITIONAL".ToLower
                    .operation = Queries.BuildSelectString_CONDITIONAL

                Case "BuildSelectString_DISTINCT".ToLower
                    .operation = Queries.BuildSelectString_DISTINCT

                Case "BuildTopString".ToLower
                    .operation = Queries.BuildTopString

                Case "BuildTopString_CONDITIONAL".ToLower
                    .operation = Queries.BuildTopString_CONDITIONAL

                Case "BuildUpdateString".ToLower
                    .operation = Queries.BuildUpdateString

                Case "BuildSelectString_BETWEEN".ToLower
                    .operation = Queries.BuildSelectString_BETWEEN

                Case "BuildSelectString_LIKE".ToLower
                    .operation = Queries.BuildSelectString_LIKE
            End Select
            .OrderByField = Content(dMax)
            Select Case Content(dOrderBy).ToLower
                Case OrderBy.ASC.ToString.ToLower
                    .OrderRecordsBy = OrderBy.ASC
                Case OrderBy.DESC.ToString.ToLower
                    .OrderRecordsBy = OrderBy.DESC
            End Select
            .SelectColumns = ListToArray(lColumns, ReturnInfo.AsArray)
            .table = Content(dTable)
            .TopRowsToSelect = CInt(Val(Content(TopRows)))
            .UpdateKeys = ListToArray(lColumns, ReturnInfo.AsArray)
            .WhereKeys = ListToArray(lWhere, ReturnInfo.AsArray)
            .WhereOperators = ListToArray(lQueryOperators, ReturnInfo.AsArray)
        End With

        Dim con_ As Object = Content(tCon)
        If IsEmpty(tCon) Then con_ = Nothing

        Dim output_ As Output
        Select Case Content(dFormatFor).ToLower
            Case Output.Desktop.ToString.ToLower
                output_ = Output.Desktop
            Case Output.Web.ToString.ToLower
                output_ = Output.Web
        End Select

        Select Case Content(dOutput)
            Case QOutput.QData.ToString
                r = QString(q_params, QOutput.QData, con_, output_)

            Case QOutput.Commit.ToString
                r = QString(q_params, QOutput.Commit, con_, output_)

            Case QOutput.Display.ToString
                r = QString(q_params, QOutput.Display, con_, output_)

            Case QOutput.QString.ToString
                r = QString(q_params, QOutput.QString, con_, output_)

            Case QOutput.QExists.ToString
                r = QString(q_params, QOutput.QExists, con_, output_)

            Case QOutput.QCount.ToString
                r = QString(q_params, QOutput.QCount, con_, output_)

            Case QOutput.QCount_Conditional.ToString
                r = QString(q_params, QOutput.QCount_Conditional, con_, output_)

            Case QOutput.BindProperty.ToString
                r = QString(q_params, QOutput.BindProperty, con_, output_)
        End Select
        If r.Length > 0 Then
            Clipboard.SetText(r)
        End If
        Return r
    End Function

#End Region
    Private Sub Form1_FormClosing(sender As Object, e As FormClosingEventArgs) Handles Me.FormClosing
        f.bye(Me, "Bye for now")
        'SaveProgramSettings()
        'e.Cancel() = True
        'CloseButton_Click(sender, e)
        '       OutTimer.Enabled = True
    End Sub
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        For Each l As Control In Me.Controls
            If TypeOf l Is Label Then
                l.BackColor = Color.Transparent
            End If
        Next

        f.say("welcome to q factory")
        'language
        'MarkLanguage()

        'theme
        'MarkTheme(selected_theme, GreenTheme, TurqoiseTheme, VelvetTheme, PurpleTheme, WhiteTheme, BrownTheme, YellowTheme, Me)

        'application feedback
        'InitialMarkVoiceFeedback(MessagePromptOnly, MessagePromptWithVoice, VoiceOnly)

        '		g.item_f = font_

        'g.item_f = font_
        dQueryFromModule.Font = font_alt
        dOutput.Font = font_alt
        Brand.Font = font_brand

        'dialog
        'g.FormatMe(Me, Nothing, LeftBorder, RightBorder, TopBorder, BottomBorder, TopLine, BottomLine, DialogTitle, SystemCloseButton, MinimizeButton, CloseButton, HelpButton, MenuStrip, SystemCloseButton, SystemCloseButton, TitleBar, FooterBar, MaximizeButton, True, False, False, True, selected_theme, Feedback, TimeTimer, TimeLabel, False, False)

        TopLine.Hide()
        BottomLine.Hide()
        DialogTitle.Hide()
        gColumns.ColumnHeadersHeight = 24
        grid.ColumnHeadersHeight = 24

        Values()
    End Sub
    Private Sub Values()
        textDataSource.Text = Environment.MachineName


        '        g.FormatDataGridView(gColumns, "", "white")
        '       g.FormatDataGridView(grid, "", "white")
        'tString.PasswordChar = "*"

        'With dSource
        '    With .Items
        '        .Add("Online")
        '        .Add("Offline")
        '    End With
        '    .Text = ReadText(source_file) ' My.Settings.source__
        'End With

        'With dGroup
        '    With .Items
        '        '.Add("My Admin")
        '        '.Add("Other")
        '        .Add("All")
        '    End With
        '    .Sorted = True
        '            .Text = "My Admin"
        'End With

        'GetDataSource()

        With dQueryFromModule
            With .Items
                .Add("BuildCountString")
                .Add("BuildCountString_CONDITIONAL")
                .Add("BuildInsertString")
                .Add("BuildMaxString")
                .Add("BuildSelectString")
                .Add("BuildSelectString_CONDITIONAL")
                .Add("BuildSelectString_DISTINCT")
                .Add("BuildTopString")
                .Add("BuildTopString_CONDITIONAL")
                .Add("BuildUpdateString")
                '                .Add("BuildUpdateString_CONDITIONAL")
                .Add("BuildSelectString_BETWEEN")
                .Add("BuildSelectString_LIKE")
                .Add("DeleteString_CONDITIONAL")
            End With
            .SelectedIndex = -1
        End With

        With dQueryOperators
            With .Items
                .Add("=")
                .Add("<>")
                .Add("<")
                .Add(">")
                .Add("<=")
                .Add(">=")
            End With
        End With

        With tOperators
            .BorderStyle = BorderStyle.None
            .ScrollBars = ScrollBars.None
            .Text = "Count"
            .Text &= vbCrLf & "Max"
            .Text &= vbCrLf & "Top"
            .Text &= vbCrLf & vbCrLf
            .Text &= "Conditional"
            .Text &= vbCrLf & "Between"
            .Text &= vbCrLf & "Like"
            .Text &= vbCrLf & "Update"
            '            .Text &= vbCrLf & vbCrLf & "Select (WHERE)"
        End With

        'Clear(dOutput)

        With dOutput
            .DataSource = Nothing
            .Items.Clear()
            .Sorted = False
            With .Items
                .Add(QOutput.QData.ToString)
                .Add(QOutput.Display.ToString)
                .Add(QOutput.Commit.ToString)
                .Add(QOutput.QExists.ToString)
                .Add(QOutput.QCount.ToString)
                .Add(QOutput.QCount_Conditional.ToString)
                .Add(QOutput.BindProperty.ToString)
                .Add(QOutput.QString.ToString)
                .Add("SQL")
            End With
            .SelectedIndex = 0
        End With

        'Clear(dFormatFor)
        With dFormatFor
            .DataSource = Nothing
            .Items.Clear()
            .Sorted = False
            With .Items
                .Add(Output.Web.ToString)
                .Add(Output.Desktop.ToString)
            End With
            .SelectedIndex = 0
        End With

        'Clear(dLikeSelect)
        With dLikeSelect
            .DataSource = Nothing
            .Items.Clear()
            .Sorted = False
            With .Items
                .Add(LIKE_SELECT.AND_.ToString.Replace("_", ""))
                .Add(LIKE_SELECT.OR_.ToString.Replace("_", ""))
            End With
            .SelectedIndex = 0
        End With
        'Clear(dOrderBy)
        With dOrderBy
            .DataSource = Nothing
            .Items.Clear()
            .Sorted = False
            With .Items
                .Add(OrderBy.DESC.ToString)
                .Add(OrderBy.ASC.ToString)
            End With
            .SelectedIndex = 1
        End With

        With TopRows
            .Maximum = 1000000
            Dim max_ = Val(ReadText(top_rows_file))
            If max_ > 1000000 Then max_ = 1000000
            .Value = max_
        End With

        textPassword.Focus()
    End Sub

    Private Sub dSource_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dSource.SelectedIndexChanged
        'If InTimer.Enabled = True Then Exit Sub
        'GetDataSource()
        'XFocusMe(sender, dGroup)
    End Sub

    Private Sub XFocusMe(sender As Control, dGroup As Control)
        dGroup.Focus()

    End Sub

    Private Sub dGroup_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dGroup.SelectedIndexChanged
        'If InTimer.Enabled = True Then Exit Sub
        'If IsEmpty(tString) Then
        '    f.say("connection string is needed. you need to provide all of data source, catalog, user id and password")
        '    Return
        'End If
        'Try
        '    GetTables() '(Content(dGroup).ToString)
        'Catch ex As Exception
        '    '            dQuery.Text = (ex.ToString)
        'End Try
        'XFocusMe(sender, dTable)
    End Sub

    Private Sub dTable_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dTable.SelectedIndexChanged
        If InTimer.Enabled = True Then Exit Sub
        Try
            GetColumns(Content(dTable))
        Catch
        End Try
        'With dQueryFromModule
        '    If .SelectedIndex < 0 Or .Text.Trim.Length < 0 Then .SelectedIndex = 4
        'End With
        Try
            dQueryFromModule.SelectedIndex = 4
        Catch
        End Try
        ClearTool_Click(sender, e)
        XFocusMe(sender, dQueryFromModule)
    End Sub

    Private Sub dQueryFromModule_SelectedIndexChanged(sender As Object, e As EventArgs) Handles dQueryFromModule.SelectedIndexChanged
        If InTimer.Enabled = True Then Exit Sub
        Dim d As ComboBox = sender
        If Content(d).ToLower.Contains("insert") Or Content(d).ToLower.Contains("update") Then
            dOutput.SelectedIndex = 2
        Else
            dOutput.SelectedIndex = 0
        End If

        XFocusMe(sender, dColumns)
    End Sub

    Private Sub bSet_Click(sender As Object, e As EventArgs) Handles bSet.Click
        '        If lColumns.Items.Count < 1 Then Exit Sub
        Clear(dQuery)

        If Content(dOutput).ToLower <> "sql" Then
            ''Write(Code(), dQuery)
            dQuery.Text = Code()
            Exit Sub
        End If


        '        Clear(dQuery)
        If Content(dQueryFromModule) = "BuildSelectString" And lWhere.Items.Count > 0 Then Exit Sub

        If Content(dQueryFromModule).ToString.Contains("CONDITIONAL") And lWhere.Items.Count < 1 Then Exit Sub

        Try
            lColumns.SelectedIndex = 0
        Catch
        End Try
        Try
            lQueryOperators.SelectedIndex = 0
        Catch
        End Try
        Dim q As String = ""
        '        Dim a As String()
        Dim l_l As New List(Of String) 'lColumns
        Dim l_m As New List(Of String)
        Dim l_r As New List(Of String)
        Dim l_l_ As String()
        Dim l_m_ As New List(Of String)
        Dim l_r_ As String()

        Dim max_ As String = Content(dMax)
        If Content(dMax).Length < 1 Then max_ = "Id"

        Try
            With lColumns
                .SelectedIndex = 0
                For i As Integer = 0 To .Items.Count - 1
                    l_l.Add(.SelectedItem)
                    Try
                        lColumns.SelectedIndex += 1
                    Catch ex As Exception
                    End Try
                Next
            End With
            l_l_ = l_l.ToArray
        Catch
        End Try

        Try
            lWhere.SelectedIndex = 0
            With lWhere
                For i As Integer = 0 To .Items.Count - 1
                    l_m.Add(.SelectedItem)
                    If Content(dQueryFromModule).ToString.Contains("CONDITIONAL") Then
                        lQueryOperators.SelectedIndex = .SelectedIndex
                        l_r.Add(lQueryOperators.SelectedItem)
                    End If
                    Try
                        lWhere.SelectedIndex += 1
                    Catch ex As Exception
                    End Try
                    If Content(dQueryFromModule).ToString.Contains("CONDITIONAL") Then
                        Try
                            lQueryOperators.SelectedIndex += 1
                        Catch ex As Exception
                        End Try
                    End If
                Next
            End With
        Catch
        End Try

        Try
            If Content(dQueryFromModule).ToString.Contains("CONDITIONAL") Then
                For i As Integer = 0 To l_m.Count - 1
                    l_m_.Add(l_m(i).ToString)
                    l_m_.Add(l_r(i).ToString)
                Next
                l_r_ = l_m_.ToArray
            Else
                For i As Integer = 0 To l_m.Count - 1
                    l_m_.Add(l_m(i).ToString)
                Next
                l_r_ = l_m_.ToArray
            End If
        Catch
        End Try

        Dim order_records As OrderBy
        Select Case Content(dOrderBy).ToLower
            Case "asc"
                order_records = OrderBy.ASC
            Case "desc"
                order_records = OrderBy.DESC
        End Select

        Dim like_select As LIKE_SELECT
        Select Case Content(dLikeSelect).ToLower
            Case "and"
                like_select = LIKE_SELECT.AND_
            Case "or"
                like_select = LIKE_SELECT.OR_
        End Select

        Try
            Select Case Content(dQueryFromModule)
                Case Query.BuildCountString.ToString
                    q = BuildCountString(Content(dTable), l_l_)
                Case Query.BuildCountString_CONDITIONAL.ToString
                    q = BuildCountString_CONDITIONAL(Content(dTable), l_r_)
                Case Query.BuildInsertString.ToString
                    q = BuildInsertString(Content(dTable), l_l_)
                Case Query.BuildMaxString.ToString
                    q = BuildMaxString(Content(dTable), max_, l_l_)
                Case Query.BuildSelectString.ToString
                    q = BuildSelectString(Content(dTable), l_l_, l_r_, max_, order_records).Replace(")", "")
                Case Query.BuildSelectString_CONDITIONAL.ToString
                    q = BuildSelectString_CONDITIONAL(Content(dTable), l_l_, l_r_, max_, order_records)
                Case Query.BuildSelectString_DISTINCT.ToString
                    q = BuildSelectString_DISTINCT(Content(dTable), l_l_, l_r_).Replace(")", "")
                Case Query.BuildTopString.ToString
                    q = BuildTopString(Content(dTable), l_l_, l_r_, Content(TopRows), max_, order_records)
                'Case Query.BuildTopString_CONDITIONAL.ToString
                '    q = BuildTopString_CONDITIONAL(Content(dTable), l_l_, l_r_, Content(TopRows), max_, order_records)
                Case Query.BuildUpdateString.ToString
                    q = BuildUpdateString(Content(dTable), l_l_, l_r_)
                Case Query.BuildSelectString_BETWEEN.ToString
                    q = BuildSelectString_BETWEEN(Content(dTable), l_l_, l_r_, max_, order_records)
                Case Query.BuildSelectString_LIKE.ToString
                    q = BuildSelectString_LIKE(Content(dTable), l_l_, l_r_, max_, order_records, like_select)
                Case Query.DeleteString_CONDITIONAL.ToString
                    q = DeleteString_CONDITIONAL(Content(dTable), l_r_)
                Case Query.BuildTopString_CONDITIONAL.ToString
                    q = BuildTopString_CONDITIONAL(Content(dTable), l_l_, l_r_, Content(TopRows), max_, order_records)
            End Select
        Catch ex As Exception

        End Try

        Dim str_ As String = Prepend(q)
        dQuery.Text = str_

        Try
            Clipboard.SetText(str_)
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dSource_KeyPress(sender As Object, e As KeyPressEventArgs) Handles dSource.KeyPress, dGroup.KeyPress, dTable.KeyPress, dQueryFromModule.KeyPress, dOutput.KeyPress, dLikeSelect.KeyPress, dFormatFor.KeyPress
        AllowNothing(e)
    End Sub

    Private Sub bAdd_Click(sender As Object, e As EventArgs) Handles bAdd.Click
        ListsIncludeItem(dColumns, lColumns)
    End Sub

    Private Sub bRemove_Click(sender As Object, e As EventArgs) Handles bRemove.Click
        ListsIncludeItem(lColumns, dColumns)
    End Sub

    Private Sub bMax_Click(sender As Object, e As EventArgs) Handles bMax.Click
    End Sub

    Private Sub bAddCondition_Click(sender As Object, e As EventArgs) Handles bAddCondition.Click
        If dWhere.Items.Count < 1 Or dWhere.SelectedIndex < 0 Then Exit Sub
        If Content(dQueryFromModule).ToLower.Contains("conditional") Then
            ListsIncludeItem(dWhere, lWhere, True)
        Else
            ListsIncludeItem(dWhere, lWhere)
        End If
        '        If Content(dQueryFromModule).ToString.Contains("CONDITIONAL") Then IncludeItem(dQueryOperators, lQueryOperators)
        If dQueryOperators.SelectedIndex < 0 Then
            '            dQueryOperators.SelectedIndex = 0
            '           IncludeItem(dQueryOperators, lQueryOperators)
            lQueryOperators.Items.Add("=")
        Else
            lQueryOperators.Items.Add(dQueryOperators.SelectedItem)
        End If
    End Sub

    Private Sub bRemoveCondition_Click(sender As Object, e As EventArgs) Handles bRemoveCondition.Click
        Try
            lQueryOperators.SelectedIndex = lWhere.SelectedIndex

            If Content(dQueryFromModule).ToLower.Contains("conditional") Then
                ListsRemoveItem(lWhere)
            Else
                ListsIncludeItem(lWhere, dWhere)
            End If

            ListsRemoveItem(lQueryOperators)
            lWhere.Focus()
        Catch
        End Try
    End Sub

    Private Sub bClearColumns_Click(sender As Object, e As EventArgs) Handles bClearColumns.Click
        Try
            With lColumns
                .SelectedIndex = 0
                For i As Integer = 0 To .Items.Count - 1
                    bRemove_Click(sender, e)
                    Try
                        .SelectedIndex += 1
                    Catch ex As Exception

                    End Try
                Next
            End With
        Catch ex As Exception

        End Try
    End Sub

    Private Sub bClearSelected_Click(sender As Object, e As EventArgs) Handles bClearSelected.Click
        Try
            With lWhere
                .SelectedIndex = 0
                For i As Integer = 0 To .Items.Count - 1
                    bRemoveCondition_Click(sender, e)
                    Try
                        .SelectedIndex += 1
                    Catch ex As Exception

                    End Try
                Next
            End With
            Clear(lQueryOperators)
        Catch
        End Try
    End Sub


    Private Sub dSource_TextChanged(sender As Object, e As EventArgs) Handles dSource.TextChanged
        SaveProgramSettings()
    End Sub

    Private Sub ClearTool_Click(sender As Object, e As EventArgs) Handles ClearTool.Click, bReset.Click
        bClearColumns_Click(sender, e)
        bClearSelected_Click(sender, e)
        'Clear(tMax)
        Clear(dQuery)
    End Sub

    Private Sub LogOutToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LogOutToolStripMenuItem.Click
        CloseButton_Click(sender, e)
    End Sub

    Private Sub bChangeSource_Click(sender As Object, e As EventArgs) Handles bChangeSource.Click
    End Sub

    Private Sub TopRows_KeyPress(sender As Object, e As KeyPressEventArgs) Handles TopRows.KeyPress
        AllowNumberOnly(e)
    End Sub

    Private Sub bAddCon_Click(sender As Object, e As EventArgs) Handles bAddCon.Click
        With tCon
            With .Items
                If .Contains(tCon.Text.Trim.ToLower) = False Then .Add(tCon.Text.Trim.ToLower)
            End With
        End With
    End Sub

    Private Sub bLoc_Click(sender As Object, e As EventArgs) Handles bLoc.Click
    End Sub

    Private Sub StartButton_Click(sender As Object, e As EventArgs) Handles StartButton.Click
        DisableControl(sender)
        GetDataSource()
        GetTables() '(Content(dGroup).ToString)
        XFocusMe(dTable, dTable)
        EnableControl(sender)
    End Sub

    Private Sub GetCatalogsButton_Click(sender As Object, e As EventArgs) Handles GetCatalogsButton.Click
        If IsEmpty({textDataSource, textUserID, textPassword}, ControlsToCheck.Any) Then
            mFeedback("Catalog, UserId and Password fields are all required!", "Some required fields are missing")
            Return
        End If

        DisableControl(sender)
        BindProperty(dCatalog, QList("select name from sys.databases", "Data Source=" & Content(textDataSource) & "\SQLEXPRESS;Persist Security Info=True;User ID=" & Content(textUserID) & ";Password=" & Content(textPassword)), False)
        EnableControl(sender)
    End Sub
End Class
