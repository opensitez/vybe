use super::helpers::run_vb;

#[test]
fn lambda_capture_matrix_outer_mutation_is_reflected() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim factor As Integer = 2
        Dim fn As Func(Of Integer, Integer) = Function(v As Integer) v * factor

        factor = 5
        Console.WriteLine(fn(3))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["15"]);
}

#[test]
fn lambda_capture_matrix_local_copy_via_for_loop_capture() {
    let out = run_vb(
        r#"
Imports System
Imports System.Collections.Generic

Module M
    Sub Main()
        Dim funcs As New List(Of Func(Of Integer))()

        For i As Integer = 1 To 3
            Dim copied As Integer = i
            funcs.Add(Function() copied * 10)
        Next

        Dim a As Integer = funcs(0)()
        Dim b As Integer = funcs(1)()
        Dim c As Integer = funcs(2)()

        Console.WriteLine(a)
        Console.WriteLine(b)
        Console.WriteLine(c)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10", "20", "30"]);
}

#[test]
fn lambda_capture_matrix_nested_lambda_chain() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim scale As Integer = 3
        Dim build As Func(Of Integer, Func(Of Integer)) = _
            Function(base As Integer) Function(x As Integer) (base + x) * scale

        Dim f As Func(Of Integer) = build(4)
        Console.WriteLine(f())
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["21"]);
}

#[test]
fn lambda_capture_matrix_action_with_shared_accumulator() {
    let out = run_vb(
        r#"
Imports System

Module M
    Sub Main()
        Dim total As Integer = 0
        Dim add As Action(Of Integer) = Sub(v As Integer)
            total += v
        End Sub

        add(4)
        add(6)

        Console.WriteLine(total)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10"]);
}

#[test]
fn lambda_capture_matrix_closure_with_query_projection() {
    let out = run_vb(
        r#"
Imports System.Linq

Module M
    Sub Main()
        Dim values() As Integer = {1, 2, 3, 4}
        Dim offset As Integer = 1

        Dim projected = values.Where(Function(v) (v + offset) Mod 2 = 0).Select(Function(v) v + offset)

        Console.WriteLine(String.Join(",", projected))
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["3,5"]);
}
