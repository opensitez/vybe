' vybe-test: vb/vb_advanced_linq_xml/linq_deferred_execution
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
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim list As New List(Of Integer) From {1, 2}
        
        Dim query = From n In list Select n
        
        list.Add(3)
        
        ' Query is evaluated here, should see 3
        __Check(CStr(query.Count()), "3")
    End Sub
End Module
