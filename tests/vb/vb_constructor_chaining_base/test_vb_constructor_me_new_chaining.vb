' vybe-test: vb/vb_constructor_chaining_base/test_vb_constructor_me_new_chaining
' origin: languages/vb/tests/vb/test_vb_constructor_chaining_base.rs

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
    Public Property Name As String
    Public Property Balance As Decimal

    Public Sub New()
        Me.New("Default", 0D)
    End Sub

    Public Sub New(name As String)
        Me.New(name, 100D)
    End Sub

    Public Sub New(name As String, balance As Decimal)
        Me.Name = name
        Me.Balance = balance
    End Sub
End Class

Module Program
    Sub Main()
        Dim a1 As New Account()
        Dim a2 As New Account("Alice")
        __Check(CStr(a1.Name & ":" & a1.Balance), "Default:0")
        __Check(CStr(a2.Name & ":" & a2.Balance), "Alice:100")
    End Sub
End Module
