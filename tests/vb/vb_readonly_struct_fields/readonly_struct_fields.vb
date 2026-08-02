' vybe-test: vb/vb_readonly_struct_fields/readonly_struct_fields
' origin: languages/vb/tests/vb/test_vb_readonly_struct_fields.rs

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

Structure Point
    Public ReadOnly X As Integer
    Public ReadOnly Y As Integer
    
    Public Sub New(xVal As Integer, yVal As Integer)
        X = xVal
        Y = yVal
    End Sub
End Structure

Module M
    Sub Main()
        Dim p As New Point(10, 20)
        __Check(CStr(p.X), "10")
        __Check(CStr(p.Y), "20")
    End Sub
End Module
