//! `using static` imports: `Math`, `Console`, `Enumerable`.
use super::helpers::run_csharp;

#[test]
fn using_static_math_allows_unqualified_sqrt() {
    assert_eq!(
        run_csharp(r#"using static System.Math;
Console.WriteLine(Sqrt(16));"#),
        &["4"]
    );
}

#[test]
fn using_static_console_allows_unqualified_writeline() {
    assert_eq!(
        run_csharp(r#"using static System.Console;
WriteLine("hello");"#),
        &["hello"]
    );
}

#[test]
fn using_static_enumerable_allows_range() {
    assert_eq!(
        run_csharp(r#"using static System.Linq.Enumerable;
Console.WriteLine(string.Join(",",Range(1,4)));"#),
        &["1,2,3,4"]
    );
}

#[test]
fn using_static_string_allows_unqualified_join() {
    assert_eq!(
        run_csharp(r#"using static System.String;
Console.WriteLine(Join("-",new[]{"a","b","c"}));"#),
        &["a-b-c"]
    );
}
