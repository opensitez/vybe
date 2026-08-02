' vybe-test: vb/vb_module_const_enums/module_const_enums
' origin: languages/vb/tests/vb/test_vb_module_const_enums.rs

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

Module Constants
    Public Const Pi As Double = 3.14159
    Public Const AppName As String = "MyApp"
    
    Public Enum Mode
        Fast = 1
        Safe = 2
    End Enum
End Module

Module M
    Sub Main()
        ' Accessing module members directly
        __Check(CStr(Pi), "3.14159")
        __Check(CStr(AppName), "MyApp")
        
        Dim m As Mode = Mode.Fast
        __Check(CStr(m), "Fast")
    End Sub
End Module
