' vybe-test: vb/vb_parse_enum_ignore_case/test_vb_enum_ulong_underlying_type_parsing
' origin: languages/vb/tests/vb/test_vb_parse_enum_ignore_case.rs

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

Enum BigEnum As ULong
    MaxVal = 18000000000000000000UL
End Enum

Module Program
    Sub Main()
        Dim e As BigEnum = CType([Enum].Parse(GetType(BigEnum), "MaxVal"), BigEnum)
        __Check(CStr(CULng(e)), "18000000000000000000")
    End Sub
End Module
