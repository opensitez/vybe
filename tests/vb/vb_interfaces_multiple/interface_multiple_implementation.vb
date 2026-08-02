' vybe-test: vb/vb_interfaces_multiple/interface_multiple_implementation
' origin: languages/vb/tests/vb/test_vb_interfaces_multiple.rs

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

Interface IReader
    Sub Read()
End Interface

Interface IWriter
    Sub Write()
End Interface

Class FileIO
    Implements IReader, IWriter
    
    Public Sub Read() Implements IReader.Read
        __Check(CStr("Reading"), "Reading")
    End Sub
    
    Public Sub Write() Implements IWriter.Write
        __Check(CStr("Writing"), "Writing")
    End Sub
End Class

Module M
    Sub Main()
        Dim io As New FileIO()
        Dim r As IReader = io
        Dim w As IWriter = io
        
        r.Read()
        w.Write()
    End Sub
End Module
