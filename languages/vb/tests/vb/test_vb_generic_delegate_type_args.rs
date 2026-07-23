use super::helpers::run_vb;

// ═══════════════════════════════════════════════════════════
// VB.NET: Generic Delegates & Action/Func Overloads Surface
// ═══════════════════════════════════════════════════════════

#[test]
fn test_vb_generic_delegate_single_type_arg() {
    let src = r#"
Delegate Function Transformer(Of T)(input As T) As T

Module Program
    Private Function DoubleInt(n As Integer) As Integer
        Return n * 2
    End Function

    Sub Main()
        Dim t As Transformer(Of Integer) = AddressOf DoubleInt
        Console.WriteLine(t(21))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["42"]);
}

#[test]
fn test_vb_generic_delegate_two_type_args() {
    let src = r#"
Delegate Function Evaluator(Of TIn, TOut)(input As TIn) As TOut

Module Program
    Sub Main()
        Dim ev As Evaluator(Of String, Integer) = Function(s) s.Length
        Console.WriteLine(ev("VisualBasic"))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["11"]);
}

#[test]
fn test_vb_generic_delegate_sub_three_type_args() {
    let src = r#"
Delegate Sub MultiAction(Of T1, T2, T3)(a As T1, b As T2, c As T3)

Module Program
    Sub Main()
        Dim act As MultiAction(Of String, Integer, Boolean) = Sub(s, i, b)
            Console.WriteLine(s & "|" & i & "|" & b)
        End Sub
        act("Data", 100, True)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Data|100|True"]);
}

#[test]
fn test_vb_generic_delegate_constraint_new() {
    let src = r#"
Delegate Function Creator(Of T As New)() As T

Class Item
    Public Tag As String = "CreatedItem"
End Class

Module Program
    Sub Main()
        Dim create As Creator(Of Item) = Function() New Item()
        Dim item = create()
        Console.WriteLine(item.Tag)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["CreatedItem"]);
}

#[test]
fn test_vb_generic_delegate_constraint_structure() {
    let src = r#"
Delegate Function ValueResetter(Of T As Structure)(val As T) As T

Module Program
    Sub Main()
        Dim resetter As ValueResetter(Of Integer) = Function(v) Nothing
        Console.WriteLine(resetter(50))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["0"]);
}

#[test]
fn test_vb_generic_delegate_constraint_class() {
    let src = r#"
Delegate Function RefValidator(Of T As Class)(item As T) As Boolean

Module Program
    Sub Main()
        Dim validator As RefValidator(Of String) = Function(s) s IsNot Nothing
        Console.WriteLine(validator("OK") & "|" & validator(Nothing))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_generic_delegate_multicast_combine() {
    let src = r#"
Imports System

Delegate Sub Logger(Of T)(item As T)

Module Program
    Private Sub Log1(s As String) : Console.WriteLine("L1: " & s) : End Sub
    Private Sub Log2(s As String) : Console.WriteLine("L2: " & s) : End Sub

    Sub Main()
        Dim l As Logger(Of String) = AddressOf Log1
        l = CType([Delegate].Combine(l, New Logger(Of String)(AddressOf Log2)), Logger(Of String))
        l("Test")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["L1: Test", "L2: Test"]);
}

#[test]
fn test_vb_generic_delegate_covariance_out() {
    let src = r#"
Delegate Function Producer(Of Out T)() As T

Module Program
    Private Function ProduceString() As String
        Return "Covariant Result"
    End Function

    Sub Main()
        Dim p As Producer(Of Object) = AddressOf ProduceString
        Console.WriteLine(p().ToString())
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Covariant Result"]);
}

#[test]
fn test_vb_generic_delegate_contravariance_in() {
    let src = r#"
Delegate Sub Consumer(Of In T)(item As T)

Module Program
    Private Sub ConsumeObject(obj As Object)
        Console.WriteLine("Contravariant: " & obj.ToString())
    End Sub

    Sub Main()
        Dim c As Consumer(Of String) = AddressOf ConsumeObject
        c("InputString")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Contravariant: InputString"]);
}

#[test]
fn test_vb_func_generic_overloads_0_to_3_args() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim f0 As Func(Of String) = Function() "F0"
        Dim f1 As Func(Of Integer, String) = Function(i) "F1_" & i
        Dim f2 As Func(Of Integer, Integer, String) = Function(i, j) "F2_" & (i + j)
        Console.WriteLine(f0() & "|" & f1(10) & "|" & f2(3, 4))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["F0|F1_10|F2_7"]);
}

#[test]
fn test_vb_action_generic_overloads_0_to_3_args() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim a0 As Action = Sub() Console.WriteLine("A0")
        Dim a1 As Action(Of String) = Sub(s) Console.WriteLine("A1_" & s)
        Dim a2 As Action(Of String, Integer) = Sub(s, i) Console.WriteLine("A2_" & s & "_" & i)
        a0()
        a1("X")
        a2("Y", 99)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["A0", "A1_X", "A2_Y_99"]);
}

#[test]
fn test_vb_predicate_generic_delegate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim pred As Predicate(Of Integer) = Function(n) n Mod 2 = 0
        Console.WriteLine(pred(4) & "|" & pred(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True|False"]);
}

#[test]
fn test_vb_comparison_generic_delegate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim comp As Comparison(Of String) = Function(s1, s2) s1.Length.CompareTo(s2.Length)
        Console.WriteLine(comp("cat", "elephant") < 0)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["True"]);
}

#[test]
fn test_vb_converter_generic_delegate() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim conv As Converter(Of Double, Integer) = Function(d) CInt(Math.Round(d))
        Console.WriteLine(conv(9.6))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["10"]);
}

#[test]
fn test_vb_generic_delegate_tuple_arguments() {
    let src = r#"
Delegate Function TupleProcessor(Of T1, T2)(pair As (Key As T1, Value As T2)) As String

Module Program
    Sub Main()
        Dim tp As TupleProcessor(Of String, Integer) = Function(pair) pair.Key & "=" & pair.Value
        Console.WriteLine(tp(("Age", 30)))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Age=30"]);
}

#[test]
fn test_vb_generic_delegate_array_argument() {
    let src = r#"
Delegate Function ArrayAggregator(Of T)(items As T()) As T

Module Program
    Sub Main()
        Dim agg As ArrayAggregator(Of Integer) = Function(arr) arr(0) + arr(1) + arr(2)
        Console.WriteLine(agg({10, 20, 30}))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["60"]);
}

#[test]
fn test_vb_generic_delegate_in_generic_class() {
    let src = r#"
Class EventSource(Of T)
    Public Delegate Sub CustomHandler(sender As Object, data As T)
    Public Event OnData As CustomHandler
    Public Sub Fire(d As T)
        RaiseEvent OnData(Me, d)
    End Sub
End Class

Module Program
    Sub Main()
        Dim src As New EventSource(Of String)()
        AddHandler src.OnData, Sub(s, data) Console.WriteLine("Event: " & data)
        src.Fire("Payload")
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Event: Payload"]);
}

#[test]
fn test_vb_generic_delegate_type_inference_lambda() {
    let src = r#"
Module Program
    Private Function Apply(Of T, R)(val As T, fn As System.Func(Of T, R)) As R
        Return fn(val)
    End Function

    Sub Main()
        Dim res = Apply(5, Function(n) "Scaled_" & (n * 10))
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Scaled_50"]);
}

#[test]
fn test_vb_generic_delegate_dynamic_invoke() {
    let src = r#"
Imports System

Delegate Function GenericCalc(Of T)(a As T, b As T) As String

Module Program
    Private Function StrCat(a As String, b As String) As String
        Return a & b
    End Function

    Sub Main()
        Dim del As [Delegate] = New GenericCalc(Of String)(AddressOf StrCat)
        Dim res = del.DynamicInvoke("Hello ", "World")
        Console.WriteLine(res)
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["Hello World"]);
}

#[test]
fn test_vb_generic_delegate_recursion_simulation() {
    let src = r#"
Imports System

Module Program
    Sub Main()
        Dim factorial As Func(Of Integer, Integer) = Nothing
        factorial = Function(n)
            If n <= 1 Then Return 1
            Return n * factorial(n - 1)
        End Function
        Console.WriteLine(factorial(5))
    End Sub
End Module
"#;
    assert_eq!(run_vb(src), vec!["120"]);
}
