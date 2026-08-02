' vybe-test: vb/vb_system_string_api_matrix/string_api_matrix_queries_and_indexing
' origin: languages/vb/tests/vb/test_vb_system_string_api_matrix.rs

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

Module M
    Sub Main()
        Dim text As String = "VBNET"
        __Check(CStr(text.Length), "5")
        __Check(CStr(text.StartsWith("VB")), "True")
        __Check(CStr(text.EndsWith("NET")), "True")
        __Check(CStr(text.Contains("BN")), "True")
        __Check(CStr(text.IndexOf("N")), "2")
        __Check(CStr(text.LastIndexOf("N")), "2")
    End Sub
End Module
