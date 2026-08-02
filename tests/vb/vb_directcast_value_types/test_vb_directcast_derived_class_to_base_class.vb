' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_derived_class_to_base_class
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

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

Class Parent
    Public Value As Integer = 50
End Class

Class SubChild
    Inherits Parent
End Class

Module Program
    Sub Main()
        Dim child As SubChild = New SubChild()
        Dim p As Parent = DirectCast(child, Parent)
        __Check(CStr(p.Value), "50")
    End Sub
End Module
