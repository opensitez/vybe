' vybe-test: vb/vb_list_remove_all_predicate/test_vb_list_remove_all_struct_items
' origin: languages/vb/tests/vb/test_vb_list_remove_all_predicate.rs

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

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim pts As New List(Of Point) From {New Point(0, 0), New Point(1, 2), New Point(0, 5)}
        pts.RemoveAll(Function(p) p.X = 0)
        __Check(CStr(pts.Count & ":" & pts(0).X & "," & pts(0).Y), "1:1,2")
    End Sub
End Module
