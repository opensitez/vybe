' vybe-test: vb/vb_tuple_deconstruct_method_overloads/test_vb_tuple_deconstruct_extension_method
' origin: languages/vb/tests/vb/test_vb_tuple_deconstruct_method_overloads.rs

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

Imports System.Runtime.CompilerServices

Class Point2D
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Class

Module PointExtensions
    <Extension()>
    Public Sub Deconstruct(p As Point2D, ByRef x As Integer, ByRef y As Integer)
        x = p.X : y = p.Y
    End Sub
End Module

Module Program
    Sub Main()
        Dim pt As New Point2D(15, 25)
        Dim x As Integer = 0
        Dim y As Integer = 0
        pt.Deconstruct(x, y)
        __Check(CStr(x & "," & y), "15,25")
    End Sub
End Module
