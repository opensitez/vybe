use super::helpers::run_vb;

#[test]
fn interlocked_atomic_matrix_increment_add_exchange() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 0

        Dim inc1 As Integer = Interlocked.Increment(value)
        Dim add5 As Integer = Interlocked.Add(value, 5)
        Dim prev As Integer = Interlocked.Exchange(value, 42)

        Console.WriteLine(inc1)
        Console.WriteLine(add5)
        Console.WriteLine(prev)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["1", "6", "6", "42"]);
}

#[test]
fn interlocked_atomic_matrix_compare_exchange_success_and_fail() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 10

        Dim ok As Integer = Interlocked.CompareExchange(value, 20, 10)
        Dim fail As Integer = Interlocked.CompareExchange(value, 30, 99)

        Console.WriteLine(ok)
        Console.WriteLine(fail)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10", "20", "20"]);
}

#[test]
fn interlocked_atomic_matrix_read_modify_repeat_pattern() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 2

        For i As Integer = 0 To 2
            Dim current As Integer = Interlocked.Increment(value)
        Next

        Dim before As Integer = value
        Dim compare As Integer = Interlocked.CompareExchange(value, 100, 5)

        Console.WriteLine(value)
        Console.WriteLine(compare)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["5", "5"]);
}
