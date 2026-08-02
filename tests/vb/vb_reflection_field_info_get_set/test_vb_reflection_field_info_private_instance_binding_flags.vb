' vybe-test: vb/vb_reflection_field_info_get_set/test_vb_reflection_field_info_private_instance_binding_flags
' origin: languages/vb/tests/vb/test_vb_reflection_field_info_get_set.rs

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

Imports System.Reflection

Class Account
    Private _balance As Double = 500.0
    Public Function GetBalance() As Double : Return _balance : End Function
End Class

Module Program
    Sub Main()
        Dim acc As New Account()
        Dim field = GetType(Account).GetField("_balance", BindingFlags.Instance Or BindingFlags.NonPublic)
        __Check(CStr(field.GetValue(acc)), "500")
        field.SetValue(acc, 1000.0)
        __Check(CStr(acc.GetBalance()), "1000")
    End Sub
End Module
