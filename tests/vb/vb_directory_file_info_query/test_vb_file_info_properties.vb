' vybe-test: vb/vb_directory_file_info_query/test_vb_file_info_properties
' origin: languages/vb/tests/vb/test_vb_directory_file_info_query.rs

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
            Dim fi As New FileInfo(tempPath)
            __Check(CStr(fi.Exists), "True")
            __Check(CStr(fi.Length), "0")
        Finally
            If File.Exists(tempPath) Then File.Delete(tempPath)
        End Try
    End Sub
End Module
