' vybe-test: vb/vb_system_dictionary_matrix/dictionary_try_get_value_contract
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
        Dim map As New Dictionary(Of String, Integer)()
        map.Add("a", 10)
        Dim value As Integer = 0
        Dim found As Boolean = map.TryGetValue("a", value)
        Dim missing As Integer = 0
        Dim missingFound As Boolean = map.TryGetValue("z", missing)
        __Check(CStr(found), "True")
        __Check(CStr(value), "10")
        __Check(CStr(missingFound), "False")
        __Check(CStr(missing), "0")
    End Sub
End Module
