' vybe-test: vb/vb_null_conditional_dictionary/null_conditional_dictionary
' origin: languages/vb/tests/vb/test_vb_null_conditional_dictionary.rs

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
        Dim dict As Dictionary(Of String, String) = Nothing
        
        ' Null conditional dictionary indexing
        Dim val = dict?("Key")
        __Check(CStr(val Is Nothing), "True")
    End Sub
End Module
