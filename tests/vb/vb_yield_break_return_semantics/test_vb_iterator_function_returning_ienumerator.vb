' vybe-test: vb/vb_yield_break_return_semantics/test_vb_iterator_function_returning_ienumerator
' origin: languages/vb/tests/vb/test_vb_yield_break_return_semantics.rs

Imports System.Collections

Module Program
    Private Iterator Function GetEnumeratorDirect() As IEnumerator
        Yield "First"
        Yield "Second"
    End Function

    Sub Main()
        Dim enumr = GetEnumeratorDirect()
        While enumr.MoveNext()
            Console.WriteLine(enumr.Current.ToString())
        End While
    End Sub
End Module
