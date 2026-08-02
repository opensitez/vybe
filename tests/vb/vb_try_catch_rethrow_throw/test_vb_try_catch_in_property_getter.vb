' vybe-test: vb/vb_try_catch_rethrow_throw/test_vb_try_catch_in_property_getter
' origin: languages/vb/tests/vb/test_vb_try_catch_rethrow_throw.rs

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

Class SafeProperty
    Public ReadOnly Property SafeValue As Integer
        Get
            Try
                Dim zero As Integer = 0
                Return 100 \ zero
            Catch ex As DivideByZeroException
                Return -1
            End Try
        End Get
    End Property
End Class

Module Program
    Sub Main()
        Dim sp As New SafeProperty()
        __Check(CStr(sp.SafeValue), "-1")
    End Sub
End Module
