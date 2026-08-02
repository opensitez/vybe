' vybe-test: vb/vb_delegates_addressof/delegate_multicast
' origin: languages/vb/tests/vb/test_vb_delegates_addressof.rs

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

Delegate Sub D()
Module M
Sub P1()
__Check(CStr("1"), "1")
End Sub
Sub P2()
__Check(CStr("2"), "2")
End Sub
Sub Main()
Dim d1 As D = AddressOf P1
Dim d2 As D = AddressOf P2
Dim d3 As D = CType([Delegate].Combine(d1, d2), D)
d3()
End Sub
End Module
