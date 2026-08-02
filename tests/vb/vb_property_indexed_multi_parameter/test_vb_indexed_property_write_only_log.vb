' vybe-test: vb/vb_property_indexed_multi_parameter/test_vb_indexed_property_write_only_log
' origin: languages/vb/tests/vb/test_vb_property_indexed_multi_parameter.rs

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

Class MemoryLog
    Private logs(2) As String
    Public WriteOnly Property LogEntry(index As Integer) As String
        Set(value As String)
            logs(index) = "LOG: " & value
        End Set
    End Property
    Public Function ReadLog(index As Integer) As String
        Return logs(index)
    End Function
End Class

Module Program
    Sub Main()
        Dim l As New MemoryLog()
        l.LogEntry(0) = "Started"
        __Check(CStr(l.ReadLog(0)), "LOG: Started")
    End Sub
End Module
