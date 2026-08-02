' vybe-test: vb/vb_advanced_linq_xml/linq_single_singleordefault
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
        Dim nums = {1}
        __Check(CStr(nums.Single()), "1")
        
        Dim nums2 = {1, 2}
        Try
            nums2.SingleOrDefault()
        Catch
            __Check(CStr("Error"), "Error")
        End Try
    End Sub
End Module
