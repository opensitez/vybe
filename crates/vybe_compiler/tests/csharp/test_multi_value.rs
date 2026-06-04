//! C# multi-value tuple returns. C# emits `return (a, b);` for tuple
//! literals and we've added `var (a, b) = f();` to the grammar — both
//! sides lower to the WASM multi-value ABI via the shared pre-scan.
//! The callee sets `chunk.result_arity = N`; the caller destructures
//! directly off the stack without a heap ValueTuple allocation.

use super::helpers::run_csharp;

#[test]
fn tuple_return_and_deconstruct() {
    let out = run_csharp(
        r#"
class Program {
    public static (int, int) Swap(int a, int b) {
        return (b, a);
    }
    public static void Run() {
        var (x, y) = Swap(1, 2);
        Console.WriteLine(x);
        Console.WriteLine(y);
    }
}
Program.Run();
"#,
    );
    assert_eq!(out, vec!["2", "1"]);
}

#[test]
fn three_value_deconstruct() {
    let out = run_csharp(
        r#"
class Program {
    public static (int, int, int) Rgb() {
        return (10, 20, 30);
    }
    public static void Run() {
        var (r, g, b) = Rgb();
        Console.WriteLine(r);
        Console.WriteLine(g);
        Console.WriteLine(b);
    }
}
Program.Run();
"#,
    );
    assert_eq!(out, vec!["10", "20", "30"]);
}
