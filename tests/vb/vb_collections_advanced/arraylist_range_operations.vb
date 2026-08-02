' vybe-test: vb/vb_collections_advanced/arraylist_range_operations
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

        ' InsertRange
        Dim ins As New ArrayList()
        ins.Add("X")
        ins.Add("Y")
        al.InsertRange(2, ins)
        __Check(CStr(al.Count), "7")
        __Check(CStr(al.Item(2)), "X")
        __Check(CStr(al.Item(3)), "Y")

        ' RemoveRange
        al.RemoveRange(2, 2)
        __Check(CStr(al.Count), "5")
        __Check(CStr(al.Item(2)), "C")

        ' GetRange
        Dim sub1 As ArrayList = al.GetRange(1, 3)
        __Check(CStr(sub1.Count), "3")
        __Check(CStr(sub1.Item(0)), "B")

        ' SetRange
        Dim rep As New ArrayList()
        rep.Add("P")
        rep.Add("Q")
        al.SetRange(1, rep)
        __Check(CStr(al.Item(1)), "P")
        __Check(CStr(al.Item(2)), "Q")

        ' Clone
        Dim cloned As ArrayList = al.Clone()
        __Check(CStr(cloned.Count), "5")
        cloned.Add("NEW")
        __Check(CStr(cloned.Count), "6")
        __Check(CStr(al.Count), "5")
    End Sub
End Module
