' vybe-test: vb/vb_generics_constraints/generic_constraint_structure
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

Module M
    ' T must be a value type
    Function GetDefault(Of T As Structure)() As T
        Dim temp As T
        Return temp
    End Function

    Sub Main()
        Dim d As Integer = GetDefault(Of Integer)()
        __Check(CStr(d), "0")
    End Sub
End Module
