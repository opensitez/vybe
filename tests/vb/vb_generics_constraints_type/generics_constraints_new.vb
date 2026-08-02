' vybe-test: vb/vb_generics_constraints_type/generics_constraints_new
' origin: languages/vb/tests/vb/test_vb_generics_constraints_type.rs

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

' Constraint: T must have a parameterless constructor
Class Factory(Of T As New)
    Public Function Create() As T
        Return New T()
    End Function
End Class

Class MyClassWithConstructor
    Public ReadOnly Value As String = "Constructed"
End Class

Module M
    Sub Main()
        Dim f As New Factory(Of MyClassWithConstructor)()
        Dim obj = f.Create()
        __Check(CStr(obj.Value), "Constructed")
    End Sub
End Module
