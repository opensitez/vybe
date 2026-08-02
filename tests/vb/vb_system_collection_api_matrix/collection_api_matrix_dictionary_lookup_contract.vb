' vybe-test: vb/vb_system_collection_api_matrix/collection_api_matrix_dictionary_lookup_contract
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
        Dim m As New Dictionary(Of String, Integer)()
        m.Add("a", 1)
        m("b") = 2

        Dim value As Integer = -1
        Dim hasC As Boolean = m.TryGetValue("c", value)
        Dim hasA As Boolean = m.TryGetValue("a", value)

        __Check(CStr(m.ContainsKey("a")), "True")
        __Check(CStr(hasC), "False")
        __Check(CStr(hasA), "True")
        __Check(CStr(value), "1")
    End Sub
End Module
