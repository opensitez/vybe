' vybe-test: vb/vb_key_value_pair_struct_usage/test_vb_key_value_pair_construction_properties
' origin: languages/vb/tests/vb/test_vb_key_value_pair_struct_usage.rs

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

Module Program
    Sub Main()
        Dim kv As New KeyValuePair(Of String, Integer)("Age", 30)
        __Check(CStr(kv.Key & "=" & kv.Value), "Age=30")
    End Sub
End Module
