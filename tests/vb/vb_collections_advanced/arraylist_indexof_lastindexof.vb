' vybe-test: vb/vb_collections_advanced/arraylist_indexof_lastindexof
' origin: languages/vb/tests/vb/test_vb_collections_advanced.rs

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

Imports System.Collections
Module M
    Sub Main()
        Dim al As New ArrayList()
        al.Add("A")
        al.Add("B")
        al.Add("C")
        al.Add("D")
        al.Add("E")
        al.Add("B")
        __Check(CStr(al.IndexOf("B")), "1")
        __Check(CStr(al.IndexOf("B", 2)), "5")
        __Check(CStr(al.LastIndexOf("B")), "5")
        __Check(CStr(al.LastIndexOf("B", 3)), "1")
    End Sub
End Module
