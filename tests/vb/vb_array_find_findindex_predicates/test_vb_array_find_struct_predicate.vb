' vybe-test: vb/vb_array_find_findindex_predicates/test_vb_array_find_struct_predicate
' origin: languages/vb/tests/vb/test_vb_array_find_findindex_predicates.rs

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

Imports System

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim pts As Point() = {New Point(1, 2), New Point(3, 4), New Point(5, 6)}
        Dim target As Point = Array.Find(pts, Function(p) p.X = 3)
        __Check(CStr(target.X & "," & target.Y), "3,4")
    End Sub
End Module
