' vybe-test: vb/vb_interface_hierarchy_multiple_inheritance/test_vb_interface_struct_implementing_multiple_interfaces
' origin: languages/vb/tests/vb/test_vb_interface_hierarchy_multiple_inheritance.rs

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

Interface IX : Function GetX() As Integer : End Interface
Interface IY : Function GetY() As Integer : End Interface

Structure Point2D
    Implements IX, IY
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer)
        Me.X = x : Me.Y = y
    End Sub
    Public Function GetX() As Integer Implements IX.GetX : Return X : End Function
    Public Function GetY() As Integer Implements IY.GetY : Return Y : End Function
End Structure

Module Program
    Sub Main()
        Dim p As New Point2D(10, 20)
        Dim ix As IX = p
        Dim iy As IY = p
        __Check(CStr(ix.GetX() & "," & iy.GetY()), "10,20")
    End Sub
End Module
