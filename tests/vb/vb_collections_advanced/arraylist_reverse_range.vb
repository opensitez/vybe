' vybe-test: vb/vb_collections_advanced/arraylist_reverse_range
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
        Dim rv As New ArrayList()
        rv.Add(1)
        rv.Add(2)
        rv.Add(3)
        rv.Add(4)
        rv.Add(5)
        rv.Reverse(1, 3)
        __Check(CStr(rv.Item(0)), "1")
        __Check(CStr(rv.Item(1)), "4")
        __Check(CStr(rv.Item(2)), "3")
        __Check(CStr(rv.Item(3)), "2")
        __Check(CStr(rv.Item(4)), "5")
    End Sub
End Module
