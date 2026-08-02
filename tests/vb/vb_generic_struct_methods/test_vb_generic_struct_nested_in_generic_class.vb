' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_nested_in_generic_class
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

Class Container(Of T)
    Public Structure Entry
        Public Key As String
        Public Value As T
        Public Sub New(k As String, v As T)
            Key = k : Value = v
        End Sub
    End Structure
End Class

Module Program
    Sub Main()
        Dim e As New Container(Of Integer).Entry("Age", 25)
        __Check(CStr(e.Key & "=" & e.Value), "Age=25")
    End Sub
End Module
