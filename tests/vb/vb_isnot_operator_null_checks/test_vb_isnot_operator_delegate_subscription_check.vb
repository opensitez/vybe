' vybe-test: vb/vb_isnot_operator_null_checks/test_vb_isnot_operator_delegate_subscription_check
' origin: languages/vb/tests/vb/test_vb_isnot_operator_null_checks.rs

Imports System

Module Program
    Sub Main()
        Dim act As Action = Sub() Console.WriteLine("Action")
        Dim nullAct As Action = Nothing
        Console.WriteLine((act IsNot Nothing) & "|" & (nullAct IsNot Nothing))
    End Sub
End Module
