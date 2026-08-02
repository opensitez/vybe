' vybe-test: vb/vb_linq_union_intersect_except_comparer/test_vb_linq_intersect_structs
' origin: languages/vb/tests/vb/test_vb_linq_union_intersect_except_comparer.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Sub Main()
        Dim set1 = {New Point(0, 0), New Point(1, 1)}
        Dim set2 = {New Point(1, 1), New Point(2, 2)}
        Dim res = set1.Intersect(set2)
        __Check(CStr(res.First().X & "," & res.First().Y), "1,1")
    End Sub
End Module
