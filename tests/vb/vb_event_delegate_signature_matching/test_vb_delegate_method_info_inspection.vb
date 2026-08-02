' vybe-test: vb/vb_event_delegate_signature_matching/test_vb_delegate_method_info_inspection
' origin: languages/vb/tests/vb/test_vb_event_delegate_signature_matching.rs

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
Imports System.Reflection

Delegate Function StringOp(input As String) As Integer

Module Program
    Private Function GetLen(s As String) As Integer
        Return s.Length
    End Function

    Sub Main()
        Dim op As StringOp = AddressOf GetLen
        Dim mi As MethodInfo = op.Method
        __Check(CStr(mi.Name & "|" & mi.ReturnType.Name), "GetLen|Int32")
    End Sub
End Module
