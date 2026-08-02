' vybe-test: vb/vb_gettype_generics/gettype_generics
' origin: languages/vb/tests/vb/test_vb_gettype_generics.rs

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

Imports System.Collections.Generic

Module M
    Sub Main()
        ' GetType for a constructed generic type
        Dim t1 As Type = GetType(List(Of String))
        __Check(CStr(t1.Name), "List`1")
        
        ' GetType for an open generic type uses (Of )
        Dim t2 As Type = GetType(List(Of ))
        __Check(CStr(t2.Name), "List`1")
        
        ' GetType for multi-parameter open generic type
        Dim t3 As Type = GetType(Dictionary(Of , ))
        __Check(CStr(t3.Name), "Dictionary`2")
    End Sub
End Module
