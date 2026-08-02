' vybe-test: vb/vb_generics_constraints/generic_constraint_new
' origin: languages/vb/tests/vb/test_vb_generics_constraints.rs

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

Class Factory(Of T As New)
    Public Function CreateInstance() As T
        Return New T()
    End Function
End Class

Class Widget
    Public Name As String = "DefaultWidget"
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of Widget)()
        Dim w As Widget = f.CreateInstance()
        __Check(CStr(w.Name), "DefaultWidget")
    End Sub
End Module
