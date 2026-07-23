use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Multicast Delegates & Invocation List
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_multicast_delegate_combine_remove() {
    let src = r#"
Public Delegate Sub NotifyHandler(msg As String)

Class Notifier
    Public Shared Log As String = ""
    Public Shared Sub Method1(m As String)
        Log &= "M1:" & m & ";"
    End Sub
    Public Shared Sub Method2(m As String)
        Log &= "M2:" & m & ";"
    End Sub
End Class

Module Program
    Sub Main()
        Dim d1 As NotifyHandler = AddressOf Notifier.Method1
        Dim d2 As NotifyHandler = AddressOf Notifier.Method2
        Dim multi As NotifyHandler = CType([Delegate].Combine(d1, d2), NotifyHandler)
        multi("Hello")

        Dim singleDel As NotifyHandler = CType([Delegate].Remove(multi, d1), NotifyHandler)
        singleDel("World")

        Console.WriteLine(Notifier.Log)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["M1:Hello;M2:Hello;M2:World;"]);
}

#[test]
fn test_vb_multicast_delegate_get_invocation_list() {
    let src = r#"
Public Delegate Sub SimpleAction()

Module Program
    Sub HandlerA()
    End Sub
    Sub HandlerB()
    End Sub

    Sub Main()
        Dim d As SimpleAction = AddressOf HandlerA
        d = CType([Delegate].Combine(d, AddressOf HandlerB), SimpleAction)
        Dim list As [Delegate]() = d.GetInvocationList()
        Console.WriteLine(list.Length)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["2"]);
}

#[test]
fn test_vb_multicast_delegate_return_value_last_wins() {
    let src = r#"
Public Delegate Function ComputeFunc() As Integer

Module Program
    Function Func1() As Integer
        Return 10
    End Function
    Function Func2() As Integer
        Return 20
    End Function

    Sub Main()
        Dim f As ComputeFunc = AddressOf Func1
        f = CType([Delegate].Combine(f, AddressOf Func2), ComputeFunc)
        Console.WriteLine(f.Invoke())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["20"]);
}

#[test]
fn test_vb_delegate_null_invocation_guard() {
    let src = r#"
Public Delegate Sub SafeNotify()

Module Program
    Sub Main()
        Dim d As SafeNotify = Nothing
        d?.Invoke()
        Console.WriteLine("Safely Executed")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Safely Executed"]);
}
