' vybe-test: vb/vb_multicast_delegate_invocation/test_vb_multicast_delegate_return_value_last_wins
' origin: languages/vb/tests/vb/test_vb_multicast_delegate_invocation.rs

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

Public Delegate Function ComputeFunc() As Integer

Module Program
    Function Func1() As Integer
        Return 10
    End Function
    Function Func2() As Integer
        Return 20
    End Function

    Sub Main()
        Dim f As ComputeFunc = AddressOf Func1
        f = CType([Delegate].Combine(f, AddressOf Func2), ComputeFunc)
        __Check(CStr(f.Invoke()), "20")
    End Sub
End Module
