' vybe-test: vb/vb_system_network_matrix/network_host_entry_is_case_insensitive_for_lookup
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
        Dim first As String = Dns.GetHostEntry("localhost").HostName
        Dim second As String = Dns.GetHostEntry("LOCALHOST").HostName
        __Check(CStr(Not String.IsNullOrWhiteSpace(first)), "True")
        __Check(CStr(first.Length > 0), "True")
        __Check(CStr(second.Length > 0), "True")
        __Check(CStr(first = second), "True")
    End Sub
End Module
