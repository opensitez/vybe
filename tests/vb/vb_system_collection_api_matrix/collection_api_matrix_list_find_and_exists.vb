' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_list_find_and_exists
' origin: languages/vb/tests/vb/test_vb_system_collection_api_matrix.rs

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

Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim values As New List(Of Integer) From {10, 20, 30, 40}
        Dim has30 As Boolean = values.Contains(30)
        Dim indexOf10 As Integer = values.IndexOf(10)
        Dim indexOf50 As Integer = values.IndexOf(50)

        __Check(CStr(has30), "True")
        __Check(CStr(indexOf10), "0")
        __Check(CStr(indexOf50), "-1")
    End Sub
End Module
