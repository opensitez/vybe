' vybe-test: vb/vb_system_reflection/system_reflection_basic
' origin: languages/vb/tests/vb/test_vb_system_reflection.rs

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

Imports System.Reflection

Class Person
    Public Property Name As String
    Public Property Age As Integer
    
    Public Sub SayHello()
        __Check(CStr("Hello"), "Person")
    End Sub
End Class

Module M
    Sub Main()
        Dim p As New Person()
        Dim t As Type = p.GetType()
        
        __Check(CStr(t.Name), "2")
        
        Dim props = t.GetProperties()
        __Check(CStr(props.Length), "Hello")
        
        Dim m = t.GetMethod("SayHello")
        If m IsNot Nothing Then
            m.Invoke(p, Nothing)
        End If
    End Sub
End Module
