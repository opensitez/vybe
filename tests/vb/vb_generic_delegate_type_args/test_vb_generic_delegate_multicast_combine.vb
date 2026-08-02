' vybe-test: vb/vb_generic_delegate_type_args/test_vb_generic_delegate_multicast_combine
' origin: languages/vb/tests/vb/test_vb_generic_delegate_type_args.rs

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

Delegate Sub Logger(Of T)(item As T)

Module Program
    Private Sub Log1(s As String) : __Check(CStr("L1: " & s), "L1: Test") : End Sub
    Private Sub Log2(s As String) : __Check(CStr("L2: " & s), "L2: Test") : End Sub

    Sub Main()
        Dim l As Logger(Of String) = AddressOf Log1
        l = CType([Delegate].Combine(l, New Logger(Of String)(AddressOf Log2)), Logger(Of String))
        l("Test")
    End Sub
End Module
