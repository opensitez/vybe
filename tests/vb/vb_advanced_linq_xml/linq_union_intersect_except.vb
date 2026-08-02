' vybe-test: vb/vb_advanced_linq_xml/linq_union_intersect_except
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
        Dim a = {1, 2, 3}
        Dim b = {3, 4, 5}
        
        Dim un = a.Union(b).Count()
        Dim int = a.Intersect(b).Count()
        Dim exc = a.Except(b).Count()
        
        __Check(CStr(un & "-" & int & "-" & exc), "5-1-2")
    End Sub
End Module
