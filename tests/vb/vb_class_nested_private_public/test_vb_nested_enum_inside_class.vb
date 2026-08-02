' vybe-test: vb/vb_class_nested_private_public/test_vb_nested_enum_inside_class
' origin: languages/vb/tests/vb/test_vb_class_nested_private_public.rs

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

Class NetworkConnection
    Public Enum State
        Disconnected
        Connecting
        Connected
    End Enum

    Public ConnectionState As State = State.Disconnected
End Class

Module Program
    Sub Main()
        Dim conn As New NetworkConnection()
        conn.ConnectionState = NetworkConnection.State.Connected
        __Check(CStr(conn.ConnectionState.ToString()), "Connected")
    End Sub
End Module
