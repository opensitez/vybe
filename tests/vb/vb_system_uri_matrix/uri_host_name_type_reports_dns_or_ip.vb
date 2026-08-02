' vybe-test: vb/vb_system_uri_matrix/uri_host_name_type_reports_dns_or_ip
' origin: languages/vb/tests/vb/test_vb_system_uri_matrix.rs

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

Module M
    Sub Main()
        Dim u1 As New Uri("https://example.com")
        Dim u2 As New Uri("https://127.0.0.1")
        __Check(CStr(u1.HostNameType.ToString()), "Dns")
        __Check(CStr(u2.HostNameType.ToString()), "IPv4")
    End Sub
End Module
