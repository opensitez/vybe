' vybe-test: vb/vb_directcast_value_types/test_vb_directcast_delegate_to_exact_delegate_type
' origin: languages/vb/tests/vb/test_vb_directcast_value_types.rs

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

Delegate Function CustomFunc(x As Integer) As Integer

Module Program
    Private Function DoubleIt(x As Integer) As Integer
        Return x * 2
    End Function

    Sub Main()
        Dim del As Object = New CustomFunc(AddressOf DoubleIt)
        Dim cf As CustomFunc = DirectCast(del, CustomFunc)
        __Check(CStr(cf(15)), "30")
    End Sub
End Module
