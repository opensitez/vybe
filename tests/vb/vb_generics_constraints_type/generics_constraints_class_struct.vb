' vybe-test: vb/vb_generics_constraints_type/generics_constraints_class_struct
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

' Constraint: T must be a reference type
Class RefContainer(Of T As Class)
    Public Item As T
End Class

' Constraint: T must be a value type
Class ValContainer(Of T As Structure)
    Public Item As T
End Class

Module M
    Sub Main()
        Dim r As New RefContainer(Of String)()
        r.Item = "Hello"
        __Check(CStr(r.Item), "Hello")
        
        Dim v As New ValContainer(Of Integer)()
        v.Item = 42
        __Check(CStr(v.Item), "42")
    End Sub
End Module
