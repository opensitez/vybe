' vybe-test: vb/vb_type_characters_decl/type_characters_declarations
' origin: languages/vb/tests/vb/test_vb_type_characters_decl.rs

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
        ' Type characters specify the variable's type without an As clause
        Dim a$ = "Hello" ' String
        Dim b% = 42      ' Integer
        Dim c& = 100000L ' Long
        Dim d! = 1.5F    ' Single
        Dim e# = 2.5     ' Double
        Dim f@ = 3.5D    ' Decimal
        
        __Check(CStr(a.GetType().Name), "String")
        __Check(CStr(b.GetType().Name), "Int32")
        __Check(CStr(c.GetType().Name), "Int64")
        __Check(CStr(d.GetType().Name), "Single")
        __Check(CStr(e.GetType().Name), "Double")
        __Check(CStr(f.GetType().Name), "Decimal")
    End Sub
End Module
