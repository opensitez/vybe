' vybe-test: vb/vb_full_domain_model_simulation/test_vb_domain_audit_log_immutable_record
' origin: languages/vb/tests/vb/test_vb_full_domain_model_simulation.rs

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

Structure AuditRecord
    Public ReadOnly Timestamp As DateTime
    Public ReadOnly Action As String
    Public ReadOnly User As String

    Public Sub New(act As String, u As String)
        Timestamp = New DateTime(2025, 1, 1)
        Action = act
        User = u
    End Sub
End Structure

Module Program
    Sub Main()
        Dim rec As New AuditRecord("LOGIN", "Admin")
        __Check(CStr(rec.User & "|" & rec.Action), "Admin|LOGIN")
    End Sub
End Module
