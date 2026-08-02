' vybe-test: vb/vb_advanced_linq_xml/linq_query_syntax_vs_method
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

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

Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        ' Mix query syntax with method call
        Dim list = (From n In nums Select n * 2).ToList()
        
        __Check(CStr(list.Count), "3")
        __Check(CStr(list(0)), "2")
    End Sub
End Module
