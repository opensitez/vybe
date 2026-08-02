' vybe-test: vb/vb_system_dictionary_matrix/dictionary_remove_returns_flags
' origin: languages/vb/tests/vb/test_vb_system_dictionary_matrix.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        Dim map As New Dictionary(Of Integer, String)()
        map.Add(1, "one")
        map.Add(2, "two")
        Dim first As Boolean = map.Remove(1)
        Dim second As Boolean = map.Remove(3)
        __Check(CStr(first), "True")
        __Check(CStr(second), "False")
        __Check(CStr(map.Count), "1")
    End Sub
End Module
