' vybe-test: vb/vb_array_convertall_transformations/test_vb_array_convertall_struct_transformation
' origin: languages/vb/tests/vb/test_vb_array_convertall_transformations.rs

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
        Dim coords As Integer() = {10, 20, 30, 40}
        ' Take pairs to points
        Dim pts As Point() = {New Point(10, 20), New Point(30, 40)}
        Dim transformed As Point() = Array.ConvertAll(pts, Function(p) New Point(p.X * 2, p.Y * 2))
        __Check(CStr(transformed(0).X & "," & transformed(0).Y), "20,40")
    End Sub
End Module
