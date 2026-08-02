' vybe-test: vb/vb_interface_default_methods_adv/test_vb_interface_multiple_interface_inheritance
' origin: languages/vb/tests/vb/test_vb_interface_default_methods_adv.rs

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

Interface IReadable
    Sub Read()
End Interface

Interface IWritable
    Sub Write()
End Interface

Interface IStreamable
    Inherits IReadable, IWritable
    Sub Flush()
End Interface

Class MemoryStreamHandler
    Implements IStreamable
    Public Sub Read() Implements IReadable.Read : __Check(CStr("Read"), "Read") : End Sub
    Public Sub Write() Implements IWritable.Write : __Check(CStr("Write"), "Write") : End Sub
    Public Sub Flush() Implements IStreamable.Flush : __Check(CStr("Flush"), "Flush") : End Sub
End Class

Module Program
    Sub Main()
        Dim s As IStreamable = New MemoryStreamHandler()
        s.Read()
        s.Write()
        s.Flush()
    End Sub
End Module
