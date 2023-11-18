'from iNovation
Imports iNovation.Code.General
Imports iNovation.Code.Desktop
Public Class Service


#Region "IO"
    Private Shared ReadOnly Property app_data_directory As String = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData) & "\iNovation Digital Works\QFactory"
    Private Shared ReadOnly Property settings_sounds_file As String = app_data_directory & "\Sounds.txt"

    Public Shared ReadOnly Property SoundsOn As Boolean
        Get
            Return Boolean.Parse(ReadText(settings_sounds_file))
        End Get
    End Property

    Public Shared Sub UpdateSoundsOn(state As Boolean)
        WriteText(settings_sounds_file, state, False, False)
    End Sub

#End Region

End Class
