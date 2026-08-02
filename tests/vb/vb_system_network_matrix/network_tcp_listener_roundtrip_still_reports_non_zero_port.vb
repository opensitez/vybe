' vybe-test: vb/vb_system_network_matrix/network_tcp_listener_roundtrip_still_reports_non_zero_port
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
Imports System.Net.Sockets

Module M
    Sub Main()
        Dim listener As New TcpListener(0)
        listener.Start()
        Dim endpoint As IPEndPoint = CType(listener.LocalEndpoint, IPEndPoint)
        __Check(CStr(endpoint.Port > 0), "True")
        listener.Stop()
        __Check(CStr(listener.LocalEndpoint IsNot Nothing), "True")
    End Sub
End Module
