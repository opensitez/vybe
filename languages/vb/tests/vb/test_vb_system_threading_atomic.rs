use super::helpers::run_vb;

#[test]
fn system_threading_atomic_exchange_and_compare_exchange() {
    let out = run_vb(
        r#"
Imports System
Imports System.Threading

Module M
    Sub Main()
        Dim value As Integer = 10
        Dim exchangeResult As Integer = Interlocked.Exchange(value, 20)
        Console.WriteLine(exchangeResult)
        Console.WriteLine(value)

        Dim failedCompare As Integer = Interlocked.CompareExchange(value, 30, 5)
        Console.WriteLine(failedCompare)
        Console.WriteLine(value)

        Dim successCompare As Integer = Interlocked.CompareExchange(value, 30, 20)
        Console.WriteLine(successCompare)
        Console.WriteLine(value)
    End Sub
End Module
"#,
    );

    assert_eq!(out, vec!["10", "20", "20", "20", "20", "30"]);
}
