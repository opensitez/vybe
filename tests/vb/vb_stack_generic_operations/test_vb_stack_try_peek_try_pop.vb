' vybe-test: vb/vb_stack_generic_operations/test_vb_stack_try_peek_try_pop
' origin: languages/vb/tests/vb/test_vb_stack_generic_operations.rs

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

Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Double)()
        Dim topVal As Double
        Dim okPeek As Boolean = st.TryPeek(topVal)
        Dim okPop As Boolean = st.TryPop(topVal)
        __Check(CStr(okPeek), "False")
        __Check(CStr(okPop), "False")

        st.Push(3.14)
        okPop = st.TryPop(topVal)
        __Check(CStr(okPop), "True")
        __Check(CStr(topVal), "3.14")
    End Sub
End Module
