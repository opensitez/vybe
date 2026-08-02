' vybe-test: vb/vb_system_guid_adv/system_guid_advanced
' origin: languages/vb/tests/vb/test_vb_system_guid_adv.rs

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
        Dim g1 As Guid = Guid.NewGuid()
        Dim g2 As Guid = Guid.NewGuid()
        
        __Check(CStr(g1.Equals(g2)), "False")
        
        Dim strGuid As String = "d87a74a4-5694-4d8b-a3ed-3085794711f1"
        Dim parsedGuid As Guid
        If Guid.TryParse(strGuid, parsedGuid) Then
            __Check(CStr("Parsed"), "Parsed")
        End If
        
        __Check(CStr(parsedGuid.ToString("D").ToLower()), "d87a74a4-5694-4d8b-a3ed-3085794711f1")
    End Sub
End Module
