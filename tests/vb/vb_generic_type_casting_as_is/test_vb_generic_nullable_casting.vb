' vybe-test: vb/vb_generic_type_casting_as_is/test_vb_generic_nullable_casting
' origin: languages/vb/tests/vb/test_vb_generic_type_casting_as_is.rs

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

Module Program
    Private Function AsNullable(Of T As Structure)(obj As Object) As Nullable(Of T)
        If TypeOf obj Is T Then
            Return CType(obj, T)
        End If
        Return Nothing
    End Function

    Sub Main()
        Dim n1 = AsNullable(Of Integer)(50)
        Dim n2 = AsNullable(Of Integer)("NotInt")
        __Check(CStr(n1.HasValue & ":" & n1.GetValueOrDefault() & "|" & n2.HasValue), "True:50|False")
    End Sub
End Module
