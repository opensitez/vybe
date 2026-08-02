' vybe-test: vb/vb_type_characters/type_characters_variables_and_literals
' origin: languages/vb/tests/vb/test_vb_type_characters.rs

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
        ' Integer type character is %
        Dim num% = 100
        
        ' Long type character is &
        Dim bigNum& = 9999999999
        
        ' Decimal type character is @
        Dim money@ = 99.99@
        
        ' Single type character is !
        Dim float! = 3.14!
        
        ' String type character is $
        Dim text$ = "Hello"
        
        __Check(CStr(num.GetType().Name), "Int32")
        __Check(CStr(bigNum.GetType().Name), "Int64")
        __Check(CStr(money.GetType().Name), "Decimal")
        __Check(CStr(float.GetType().Name), "Single")
        __Check(CStr(text.GetType().Name), "String")
    End Sub
End Module
