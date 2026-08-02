' vybe-test: vb/vb_dictionary_operations/dictionary_try_get_value_returns_false_and_default_on_miss
' origin: languages/vb/tests/vb/test_vb_dictionary_operations.rs

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
        Dim map As New Dictionary(Of String, Integer)()
        Dim value As Integer
        __Check(CStr(map.TryGetValue("nope", value)), "False")
        __Check(CStr(value), "0")
    End Sub
End Module
