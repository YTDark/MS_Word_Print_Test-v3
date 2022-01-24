Module Module_Print


    Private _MSWord As MSwordEx.MSO.MSWord
    Public Property MSWord() As MSwordEx.MSO.MSWord
        Get
            If _MSWord Is Nothing OrElse _MSWord.AppIsDead Then
                _MSWord = New MSwordEx.MSO.MSWord(My.Forms.Form1)
            End If
            Return _MSWord
        End Get
        Set(ByVal value As MSwordEx.MSO.MSWord)
            _MSWord = value
        End Set
    End Property


    Public TempDir As String = My.Computer.FileSystem.SpecialDirectories.Temp & "\Abdulla_Aldosari_Test_Print\"


    Public Function GetNewFileName(ByVal FileNameStartWith As String, ByVal Extension As String) As String
        If Not IO.Directory.Exists(TempDir) Then
            IO.Directory.CreateDirectory(TempDir)
        End If
        Return TempDir & FileNameStartWith & Guid.NewGuid().ToString("N") & "." & Extension
    End Function





End Module
