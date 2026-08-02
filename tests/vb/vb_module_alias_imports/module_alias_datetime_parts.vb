' vybe-test: vb/vb_module_alias_imports/module_alias_datetime_parts
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

Imports Dates = System

Module M
    Sub Main()
        Dim now As Dates.DateTime = Dates.DateTime.Parse("2026-02-03T00:00:00")
        __Check(CStr(now.Month), "2")
        __Check(CStr(CStr(now.DayOfWeek)), "Tuesday")
    End Sub
End Module
