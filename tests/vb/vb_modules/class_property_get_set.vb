' vybe-test: vb/vb_modules/class_property_get_set
' origin: languages/vb/tests/vb/test_vb_modules.rs

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

Class Account
    Private _balance As Integer = 0
    Public Property Balance As Integer
        Get
            Return _balance
        End Get
        Set(value As Integer)
            If value >= 0 Then _balance = value
        End Set
    End Property
End Class

Module M
    Sub Main()
        Dim a As New Account()
        a.Balance = 100
        __Check(CStr(a.Balance), "100")
        a.Balance = -50
        __Check(CStr(a.Balance), "100")
    End Sub
End Module
