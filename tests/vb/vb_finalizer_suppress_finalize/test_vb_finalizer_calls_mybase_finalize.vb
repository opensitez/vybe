' vybe-test: vb/vb_finalizer_suppress_finalize/test_vb_finalizer_calls_mybase_finalize
' origin: languages/vb/tests/vb/test_vb_finalizer_suppress_finalize.rs

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

Class BaseFinalizer
    Public Shared BaseRan As Boolean = False
    Protected Overrides Sub Finalize()
        BaseRan = True
    End Sub
End Class

Class DerivedFinalizer
    Inherits BaseFinalizer
    Public Shared DerivedRan As Boolean = False
    Protected Overrides Sub Finalize()
        DerivedRan = True
        MyBase.Finalize()
    End Sub
End Class

Module Program
    Sub Main()
        Sub()
            Dim d As New DerivedFinalizer()
        End Sub()

        GC.Collect()
        GC.WaitForPendingFinalizers()

        __Check(CStr(DerivedFinalizer.DerivedRan & "|" & BaseFinalizer.BaseRan), "True|True")
    End Sub
End Module
