' vybe-test: vb/vb_type_characters_adv/type_characters_variable_declaration
' origin: languages/vb/tests/vb/test_vb_type_characters_adv.rs

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

Module M
    Sub Main()
        ' Type characters define the type without explicit 'As Type'
        Dim i% = 10     ' Integer
        Dim l& = 100    ' Long
        Dim d@ = 10.5D  ' Decimal
        Dim s! = 2.5!   ' Single
        Dim f# = 3.14#  ' Double
        Dim str$ = "VB" ' String
        
        __Check(CStr(i.GetType().Name), "Int32")
        __Check(CStr(l.GetType().Name), "Int64")
        __Check(CStr(d.GetType().Name), "Decimal")
        __Check(CStr(s.GetType().Name), "Single")
        __Check(CStr(f.GetType().Name), "Double")
        __Check(CStr(str.GetType().Name), "String")
    End Sub
End Module
