' vybe-test: vb/vb_system_cast_runtime_matrix/cast_runtime_directcast_downcast_success
' origin: languages/vb/tests/vb/test_vb_system_cast_runtime_matrix.rs

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

Class Base
End Class

Class Child
    Inherits Base
    Public Property Name As String = "child"
End Class

Module M
    Sub Main()
        Dim b As Base = New Child()
        Dim c As Child = DirectCast(b, Child)

        __Check(CStr(c.Name), "child")
    End Sub
End Module
