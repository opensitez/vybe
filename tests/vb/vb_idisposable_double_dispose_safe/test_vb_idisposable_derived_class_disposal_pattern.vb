' vybe-test: vb/vb_idisposable_double_dispose_safe/test_vb_idisposable_derived_class_disposal_pattern
' origin: languages/vb/tests/vb/test_vb_idisposable_double_dispose_safe.rs

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

Class BaseRes
    Implements IDisposable
    Protected BaseDisposed As Boolean = False

    Protected Overridable Sub Dispose(disposing As Boolean)
        If Not BaseDisposed Then
            If disposing Then __Check(CStr("Base Managed Cleaned"), "Derived Managed Cleaned")
            BaseDisposed = True
        End If
    End Sub

    Public Sub Dispose() Implements IDisposable.Dispose
        Dispose(True)
        GC.SuppressFinalize(Me)
    End Sub
End Class

Class DerivedRes
    Inherits BaseRes

    Private DerivedDisposed As Boolean = False

    Protected Overrides Sub Dispose(disposing As Boolean)
        If Not DerivedDisposed Then
            If disposing Then __Check(CStr("Derived Managed Cleaned"), "Base Managed Cleaned")
            DerivedDisposed = True
        End If
        MyBase.Dispose(disposing)
    End Sub
End Class

Module Program
    Sub Main()
        Dim res As New DerivedRes()
        res.Dispose()
    End Sub
End Module
