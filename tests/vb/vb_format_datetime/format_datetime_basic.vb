' vybe-test: vb/vb_format_datetime/format_datetime_basic
' origin: languages/vb/tests/vb/test_vb_format_datetime.rs

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
        Dim dt As Date = #1/1/2026 14:30:00#
        
        Dim f1 As String = FormatDateTime(dt, DateFormat.GeneralDate)
        Dim f2 As String = FormatDateTime(dt, DateFormat.LongDate)
        Dim f3 As String = FormatDateTime(dt, DateFormat.ShortDate)
        Dim f4 As String = FormatDateTime(dt, DateFormat.LongTime)
        Dim f5 As String = FormatDateTime(dt, DateFormat.ShortTime)
        
        __Check(CStr("FormattedDates"), "FormattedDates")
    End Sub
End Module
