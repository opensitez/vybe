' vybe-test: vb/vb_class/class_tostring_method
' origin: languages/vb/tests/vb/vb_class_test.rs

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

Module Program
    Class Point
        Public X As Integer
        Public Y As Integer

        Sub New(x As Integer, y As Integer)
            Me.X = x
            Me.Y = y
        End Sub

        Function ToString() As String
            ToString = "(" & CStr(Me.X) & ", " & CStr(Me.Y) & ")"
        End Function
    End Class

    Sub Main()
        Dim p As New Point(10, 20)
        __Check(CStr(p.ToString()), "(10, 20)")
    End Sub
End Module
