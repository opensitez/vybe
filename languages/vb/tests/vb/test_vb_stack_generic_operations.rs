use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Stack(Of T) Operations & Semantics
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_stack_push_pop_lifo_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Integer)()
        st.Push(10)
        st.Push(20)
        st.Push(30)
        Console.WriteLine(st.Pop())
        Console.WriteLine(st.Pop())
        Console.WriteLine(st.Pop())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["30", "20", "10"]);
}

#[test]
fn test_vb_stack_peek_non_destructive() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of String)()
        st.Push("Top")
        Console.WriteLine(st.Peek())
        Console.WriteLine(st.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Top", "1"]);
}

#[test]
fn test_vb_stack_try_peek_try_pop() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Double)()
        Dim topVal As Double
        Dim okPeek As Boolean = st.TryPeek(topVal)
        Dim okPop As Boolean = st.TryPop(topVal)
        Console.WriteLine(okPeek)
        Console.WriteLine(okPop)

        st.Push(3.14)
        okPop = st.TryPop(topVal)
        Console.WriteLine(okPop)
        Console.WriteLine(topVal)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["False", "False", "True", "3.14"]);
}

#[test]
fn test_vb_stack_contains_value() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of String)()
        st.Push("Alpha")
        st.Push("Beta")
        Console.WriteLine(st.Contains("Alpha"))
        Console.WriteLine(st.Contains("Gamma"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True", "False"]);
}

#[test]
fn test_vb_stack_to_array_order() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Integer)()
        st.Push(1)
        st.Push(2)
        st.Push(3)
        Dim arr As Integer() = st.ToArray()
        Console.WriteLine(String.Join(",", arr))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["3,2,1"]);
}

#[test]
fn test_vb_stack_trim_excess() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Integer)(100)
        st.Push(1)
        st.TrimExcess()
        Console.WriteLine(st.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["1"]);
}

#[test]
fn test_vb_stack_clear() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of Integer)()
        st.Push(1)
        st.Clear()
        Console.WriteLine(st.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_stack_enumeration_without_popping() {
    let src = r#"
Imports System.Collections.Generic

Module Program
    Sub Main()
        Dim st As New Stack(Of String)()
        st.Push("A")
        st.Push("B")
        Console.WriteLine(String.Join(",", st))
        Console.WriteLine(st.Count)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["B,A", "2"]);
}
