' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_ref_returns_simulation
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

Structure CounterStruct(Of T)
    Public Value As T
    Public Sub Increment(byRefTarget As ByRefHolder(Of T), incrementFunc As System.Func(Of T, T))
        byRefTarget.Value = incrementFunc(byRefTarget.Value)
    End Sub
End Structure

Class ByRefHolder(Of T)
    Public Property Value As T
End Class

Module Program
    Sub Main()
        Dim cs As New CounterStruct(Of Integer)()
        Dim holder As New ByRefHolder(Of Integer)() With {.Value = 10}
        cs.Increment(holder, Function(n) n + 5)
        __Check(CStr(holder.Value), "15")
    End Sub
End Module
