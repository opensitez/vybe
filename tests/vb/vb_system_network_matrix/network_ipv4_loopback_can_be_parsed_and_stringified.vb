' vybe-test: vb/vb_system_network_matrix/network_ipv4_loopback_can_be_parsed_and_stringified
' origin: languages/vb/tests/vb/test_vb_system_network_matrix.rs

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

Imports System.Net

Module M
    Sub Main()
        Dim ip As IPAddress = IPAddress.Parse("127.0.0.1")
        __Check(CStr(ip.ToString()), "127.0.0.1")
        __Check(CStr(ip.AddressFamily.ToString()), "InterNetwork")
    End Sub
End Module
