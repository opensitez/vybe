' vybe-test: vb/vb_linq_select_many_projections/test_vb_linq_select_many_with_index
' origin: languages/vb/tests/vb/test_vb_linq_select_many_projections.rs

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

Module Program
    Sub Main()
        Dim sentences = {"Hello World", "VB NET"}
        Dim result = sentences.SelectMany(Function(s, idx) s.Split(" "c).Select(Function(w) idx & ":" & w))
        __Check(CStr(String.Join(",", result)), "0:Hello,0:World,1:VB,1:NET")
    End Sub
End Module
