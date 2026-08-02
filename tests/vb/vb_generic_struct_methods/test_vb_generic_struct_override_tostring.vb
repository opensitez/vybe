' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_override_tostring
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

Structure Vector2D(Of T)
    Public X As T
    Public Y As T
    Public Sub New(x As T, y As T)
        Me.X = x : Me.Y = y
    End Sub
    Public Overrides Function ToString() As String
        Return "[" & X.ToString() & ", " & Y.ToString() & "]"
    End Function
End Structure

Module Program
    Sub Main()
        Dim v As New Vector2D(Of Single)(1.5F, 2.5F)
        __Check(CStr(v.ToString()), "[1.5, 2.5]")
    End Sub
End Module
