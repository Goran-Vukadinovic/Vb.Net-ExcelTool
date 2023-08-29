Imports System.IO

Module Logging
    Public Function WriteLogging(LogginInfo As String, fileapath As String)
        File.AppendAllText(fileapath, LogginInfo + Environment.NewLine)
    End Function

End Module
