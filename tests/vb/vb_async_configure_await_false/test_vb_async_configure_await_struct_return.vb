' vybe-test: vb/vb_async_configure_await_false/test_vb_async_configure_await_struct_return
' origin: languages/vb/tests/vb/test_vb_async_configure_await_false.rs

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

Imports System.Threading.Tasks

Structure Point
    Public X As Integer
    Public Y As Integer
    Public Sub New(x As Integer, y As Integer) : Me.X = x : Me.Y = y : End Sub
End Structure

Module Program
    Private Async Function GetPointAsync() As Task(Of Point)
        Await Task.Delay(5).ConfigureAwait(False)
        Return New Point(100, 200)
    End Function

    Sub Main()
        Dim t = GetPointAsync()
        __Check(CStr(t.Result.X & "," & t.Result.Y), "100,200")
    End Sub
End Module
