' vybe-test: vb/vb_generic_struct_methods/test_vb_generic_struct_nullable_field
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

Imports System

Structure NullableHolder(Of T As Structure)
    Public Value As Nullable(Of T)
    Public Sub New(v As T)
        Value = v
    End Sub
End Structure

Module Program
    Sub Main()
        Dim nh As New NullableHolder(Of Integer)(55)
        __Check(CStr(nh.Value.HasValue & ":" & nh.Value.Value), "True:55")
    End Sub
End Module
