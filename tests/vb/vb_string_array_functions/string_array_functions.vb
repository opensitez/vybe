' vybe-test: vb/vb_string_array_functions/string_array_functions
' origin: languages/vb/tests/vb/test_vb_string_array_functions.rs

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
        Dim csv As String = "A,B,C"
        
        ' Split string into array
        Dim arr() As String = Split(csv, ",")
        __Check(CStr(arr.Length), "3")
        __Check(CStr(arr(1)), "B")
        
        ' Join array into string
        Dim joined = Join(arr, "-")
        __Check(CStr(joined), "A-B-C")
        
        ' Filter array
        Dim words() As String = {"Apple", "Banana", "Cherry", "Apricot"}
        Dim filtered = Filter(words, "Ap")
        __Check(CStr(filtered.Length), "2")
        __Check(CStr(filtered(0)), "Apple")
        __Check(CStr(filtered(1)), "Apricot")
    End Sub
End Module
