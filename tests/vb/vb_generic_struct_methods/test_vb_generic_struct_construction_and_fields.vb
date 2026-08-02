' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_construction_and_fields
' origin: languages/vb/tests/vb/test_vb_generic_struct_methods.rs

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

Structure Point(Of T)
    Public X As T
    Public Y As T
    Public Sub New(x As T, y As T)
        Me.X = x : Me.Y = y
    End Sub
End Structure

Module Program
    Sub Main()
        Dim p As New Point(Of Integer)(10, 20)
        __Check(CStr(p.X & "," & p.Y), "10,20")
    End Sub
End Module
