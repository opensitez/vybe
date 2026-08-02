' vybe-test: vb/vb_property_protected_internal_mix/test_vb_property_public_get_protected_set
' origin: languages/vb/tests/vb/test_vb_property_protected_internal_mix.rs

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

Class BaseAccount
    Public Property Balance As Decimal { Get; Protected Set; }
    Public Sub New(bal As Decimal)
        Balance = bal
    End Sub
End Class

Class CheckingAccount
    Inherits BaseAccount
    Public Sub New(bal As Decimal)
        MyBase.New(bal)
    End Sub
    Public Sub Deposit(amount As Decimal)
        Balance += amount
    End Sub
End Class

Module Program
    Sub Main()
        Dim acc As New CheckingAccount(100D)
        acc.Deposit(50D)
        __Check(CStr(acc.Balance), "150")
    End Sub
End Module
