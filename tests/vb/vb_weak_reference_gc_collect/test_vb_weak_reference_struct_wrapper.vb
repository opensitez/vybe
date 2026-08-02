' vybe-test: vb/vb_weak_reference_gc_collect/test_vb_weak_reference_struct_wrapper
' origin: languages/vb/tests/vb/test_vb_weak_reference_gc_collect.rs

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

Imports System

Structure ManagedWeakHandle(Of T As Class)
    Private ref As WeakReference(Of T)
    Public Sub New(target As T)
        ref = New WeakReference(Of T)(target)
    End Sub
    Public ReadOnly Property Value As T
        Get
            Dim target As T = Nothing
            If ref IsNot Nothing AndAlso ref.TryGetTarget(target) Then Return target
            Return Nothing
        End Get
    End Property
End Structure

Class Node
    Public Name As String = "StructWrappedNode"
End Class

Module Program
    Sub Main()
        Dim n As New Node()
        Dim handle As New ManagedWeakHandle(Of Node)(n)
        __Check(CStr(handle.Value.Name), "StructWrappedNode")
    End Sub
End Module
