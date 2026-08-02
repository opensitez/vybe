' vybe-test: vb/vb_file_stream_read_write/test_vb_file_write_read_lines
' origin: languages/vb/tests/vb/test_vb_file_stream_read_write.rs

' Vybe test harness — Visual Basic.
'
' Real VB source alongside harness/go/check.go and harness/js/check.js, the way
' test262's assert.js is JavaScript.
'
' A test's verdict is its EXIT CODE. __Check prints its diagnostic BEFORE
' throwing: an uncaught exception surfaces as `RuntimeError: [object]`, which
' says nothing at all.

Module VybeCheck
    Sub __Check(got As String, want As String)
        If got <> want Then
            Console.WriteLine("FAIL: want [" & want & "] got [" & got & "]")
            Throw New Exception("assertion failed")
        End If
    End Sub
End Module

Imports System.IO

Module Program
    Sub Main()
        Dim tempPath As String = Path.GetTempFileName()
        Try
            File.WriteAllLines(tempPath, New String() {"LineA", "LineB"})
            Dim lines As String() = File.ReadAllLines(tempPath)
            __Check(CStr(lines.Length), "2")
            __Check(CStr(lines(0) & "," & lines(1)), "LineA,LineB")
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub
End Module
