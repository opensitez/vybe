' vybe-test: vb/vb_gchandle_alloc_free/test_vb_gchandle_alloc_normal_and_free
' origin: languages/vb/tests/vb/test_vb_gchandle_alloc_free.rs

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

Imports System.Runtime.InteropServices

Class TargetObj
    Public Value As Integer = 42
End Class

Module Program
    Sub Main()
        Dim obj As New TargetObj()
        Dim handle As GCHandle = GCHandle.Alloc(obj, GCHandleType.Normal)
        Dim retrieved As TargetObj = CType(handle.Target, TargetObj)
        __Check(CStr(handle.IsAllocated & "|" & retrieved.Value), "True|42")
        handle.Free()
        __Check(CStr(handle.IsAllocated), "False")
    End Sub
End Module
