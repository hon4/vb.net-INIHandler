Imports System.IO
'INIHandler v1.0 by hon
'Coded: D02-M03-Y2025
Public Class INIHandler

    Public Shared Function ReadINI(ByVal INIPath As String) As Dictionary(Of String, Object)
        Dim ret As New Dictionary(Of String, Object)

        Try
            Using sr As New StreamReader(INIPath)
                Dim ln As String = sr.ReadLine()
                Dim currSect As String = Nothing
                While ln IsNot Nothing
                    'Start Decoder
                    ln = ln.Trim()
                    If ln.StartsWith("["c) AndAlso ln.EndsWith("]"c) Then
                        currSect = ln.Substring(1, ln.Length - 2).Trim()
                        ret(currSect) = New Dictionary(Of String, Object)
                    ElseIf Not ln.StartsWith(";"c) AndAlso Not String.IsNullOrEmpty(ln) Then
                        Dim key As String = ln.Substring(0, ln.IndexOf("="c))
                        If currSect Is Nothing Then
                            ret(key) = ln.Substring(ln.IndexOf("="c) + 1)
                        Else
                            ret(currSect)(key) = ln.Substring(ln.IndexOf("="c) + 1)
                        End If
                    End If
                    'End Decoder
                    ln = sr.ReadLine()
                End While
            End Using
        Catch ex As Exception
            'Just there is not INI file.
        End Try

        Return ret
    End Function

    Public Shared Sub SaveINI(ByVal INIPath As String, ByVal dict As Dictionary(Of String, Object))
        SaveINIF(INIPath, dict)
    End Sub

    Public Shared Function SaveINIF(ByVal INIPath As String, ByVal dict As Dictionary(Of String, Object)) As Boolean
        Dim out As String = ""
        For Each kvp As KeyValuePair(Of String, Object) In dict
            If kvp.Value.GetType Is GetType(Dictionary(Of String, Object)) Then
                out &= "[" & kvp.Key & "]" & vbNewLine
                For Each ikvp As KeyValuePair(Of String, Object) In kvp.Value
                    out &= ikvp.Key & "=" & ikvp.Value & vbNewLine
                Next
            Else
                out &= kvp.Key & "=" & kvp.Value & vbNewLine
            End If
        Next

        Try
            File.WriteAllText(INIPath, out)
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function
End Class
