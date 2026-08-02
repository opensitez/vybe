' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_tuple_field
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

Structure TupleHolder(Of T)
    Public Data As (Key As String, Value As T)
    Public Sub New(k As String, v As T)
        Data = (k, v)
    End Sub
End Structure

Module Program
    Sub Main()
        Dim th As New TupleHolder(Of Double)("PI", 3.14159)
        __Check(CStr(th.Data.Key & "=" & th.Data.Value), "PI=3.14159")
    End Sub
End Module
