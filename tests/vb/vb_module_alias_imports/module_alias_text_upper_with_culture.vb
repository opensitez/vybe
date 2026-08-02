' vybe-test: vb/vb_module_alias_imports/module_alias_text_upper_with_culture
' origin: languages/vb/tests/vb/test_vb_module_alias_imports.rs

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

Imports Texts = System.Text
Imports Culture = System.Globalization.CultureInfo

Module M
    Sub Main()
        Dim title As New Texts.StringBuilder("ß")
        Dim culture = Culture.InvariantCulture
        __Check(CStr(title.ToString().ToUpper(culture)), "SS")
    End Sub
End Module
