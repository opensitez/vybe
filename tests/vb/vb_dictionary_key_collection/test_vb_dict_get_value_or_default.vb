' vybe-test: vb/vb_dictionary_key_collection/test_vb_dict_get_value_or_default
' origin: languages/vb/tests/vb/test_vb_dictionary_key_collection.rs

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
        Dim dict As New Dictionary(Of String, Integer) From {{"Existing", 50}}
        Dim val1 As Integer
        dict.TryGetValue("Existing", val1)
        Dim val2 As Integer
        dict.TryGetValue("Missing", val2)
        __Check(CStr(val1), "50")
        __Check(CStr(val2), "0")
    End Sub
End Module
