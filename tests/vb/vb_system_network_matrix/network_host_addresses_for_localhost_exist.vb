' vybe-test: vb/vb_system_network_matrix/network_host_addresses_for_localhost_exist
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
        Dim addresses() As IPAddress = Dns.GetHostAddresses("localhost")
        __Check(CStr(addresses.Length >= 1), "True")
        __Check(CStr(addresses(0).ToString().Length > 0), "True")
    End Sub
End Module
