' vybe-test: vb/vb_advanced_linq_xml/linq_last_lastordefault
' origin: languages/vb/tests/vb/test_vb_advanced_linq_xml.rs

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

Imports System.Linq

Module M
    Sub Main()
        Dim nums = {1, 2, 3}
        __Check(CStr(nums.Last()), "3")
        
        Dim empty As Integer() = {}
        __Check(CStr(empty.LastOrDefault()), "0")
    End Sub
End Module
