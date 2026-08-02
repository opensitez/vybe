' vybe-test: vb/vb_iif_ternary_evaluation/test_vb_iif_eager_evaluation_both_branches
' origin: languages/vb/tests/vb/test_vb_iif_ternary_evaluation.rs

Imports Microsoft.VisualBasic

Module Program
    Private Function SideEffect(msg As String, val As Integer) As Integer
        Console.WriteLine("Effect:" & msg)
        Return val
    End Function

    Sub Main()
        ' IIf eagerly evaluates both truepart and falsepart!
        Dim res = IIf(True, SideEffect("TrueBranch", 10), SideEffect("FalseBranch", 20))
        Console.WriteLine(res)
    End Sub
End Module
