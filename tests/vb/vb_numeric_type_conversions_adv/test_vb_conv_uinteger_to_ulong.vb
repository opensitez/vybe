' vybe-test: vb/vb_numeric_type_conversions_adv/test_vb_conv_uinteger_to_ulong
' origin: languages/vb/tests/vb/test_vb_numeric_type_conversions_adv.rs

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

Module Program
    Sub Main()
        Dim ui As UInteger = 4294967295UI
        Dim ul As ULong = CULng(ui)
        __Check(CStr(ul), "4294967295")
    End Sub
End Module
