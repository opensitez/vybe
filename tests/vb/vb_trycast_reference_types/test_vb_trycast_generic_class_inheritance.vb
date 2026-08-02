' vybe-test: vb/vb_trycast_reference_types/test_vb_trycast_generic_class_inheritance
' origin: languages/vb/tests/vb/test_vb_trycast_reference_types.rs

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

Class BaseContainer(Of T)
End Class

Class DerivedContainer(Of T)
    Inherits BaseContainer(Of T)
End Class

Module Program
    Sub Main()
        Dim baseObj As BaseContainer(Of Integer) = New DerivedContainer(Of Integer)()
        Dim derivedObj As DerivedContainer(Of Integer) = TryCast(baseObj, DerivedContainer(Of Integer))
        __Check(CStr(derivedObj IsNot Nothing), "True")
    End Sub
End Module
