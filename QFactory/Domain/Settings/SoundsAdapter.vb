Imports QFactory.Service
Public Class SoundsAdapter
    Private Shared f As New iNovation.Code.Feedback()

    Public Shared Sub say(message As String, Optional async As Boolean = True)
        If Not SoundsOn Then Return
        Try
            f.say(message, async)

        Catch ex As Exception

        End Try

    End Sub
    Public Shared Sub welcome(message As String, Optional sound As String = Nothing)
        If Not SoundsOn Then Return
        Try
            f.welcome(message, sound)

        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub bye(dialog As Form, str As String, Optional dont_use_voice_feedback As Boolean = False)
        Try
            f.bye(dialog, If(SoundsOn, str, ""), dont_use_voice_feedback)

        Catch ex As Exception

        End Try
    End Sub
    Public Shared Sub bye(dialog As Form, str As String, sound_ As String, Optional dont_use_voice_feedback As Boolean = False)
        Try
            f.bye(dialog, If(SoundsOn, str, ""), sound_, dont_use_voice_feedback)

        Catch ex As Exception

        End Try
    End Sub




End Class
