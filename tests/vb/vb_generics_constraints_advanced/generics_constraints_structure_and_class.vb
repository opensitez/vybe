' vybe-test: vb/vb_generics_constraints_advanced/generics_constraints_structure_and_class
' origin: languages/vb/tests/vb/test_vb_generics_constraints_advanced.rs

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

' Structure constraint requires value type (excluding Nullable)
' Class constraint requires reference type
Class HolderRef(Of T As Class)
    Public Property Value As T
End Class

Class HolderVal(Of T As Structure)
    Public Property Value As T
End Class

Class Item
End Class

Module M
    Sub Main()
        Dim hr As New HolderRef(Of Item)()
        Dim hv As New HolderVal(Of Integer)()
        
        hr.Value = Nothing
        hv.Value = 42
        
        __Check(CStr(hr.Value Is Nothing), "True")
        __Check(CStr(hv.Value), "42")
    End Sub
End Module
