' vybe-test: vb/vb_system_culture_matrix/cultureinfo_neutral_and_specific_relationship
' origin: languages/vb/tests/vb/test_vb_system_culture_matrix.rs

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
Imports System.Globalization

Module M
    Sub Main()
        Dim specific As CultureInfo = CultureInfo.GetCultureInfo("en-US")
        Dim parentName As String = specific.Parent.Name

        __Check(CStr(specific.IsNeutralCulture), "False")
        __Check(CStr(parentName = "en"), "True")
        __Check(CStr(specific.Equals(CultureInfo.GetCultureInfo("en-US"))), "True")
    End Sub
End Module
