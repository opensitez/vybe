' vybe-test: vb/vb_convert_change_type_reflection/test_vb_convert_change_type_nullable_underlying_type
' origin: languages/vb/tests/vb/test_vb_convert_change_type_reflection.rs

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
    Private Function ChangeTypeToNullable(val As Object, targetType As Type) As Object
        Dim underlying = Nullable.GetUnderlyingType(targetType)
        Dim effectiveType = If(underlying, targetType)
        If val Is Nothing Then Return Nothing
        Return Convert.ChangeType(val, effectiveType)
    End Function

    Sub Main()
        Dim res = ChangeTypeToNullable("99", GetType(Integer?))
        __Check(CStr(res.GetType().Name & ":" & res), "Int32:99")
    End Sub
End Module
